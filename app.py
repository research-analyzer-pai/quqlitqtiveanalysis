import streamlit as st
import pandas as pd
import sqlite3, json, os, re, hashlib, uuid
from pathlib import Path
from datetime import datetime, date
from io import BytesIO
from docx import Document

ROOT=Path(__file__).resolve().parent
DB=ROOT/"q_analyst.db"
UPLOAD=ROOT/"uploads"; UPLOAD.mkdir(exist_ok=True)
st.set_page_config(page_title="Q-ANALYST",page_icon="🔎",layout="wide")

def conn():
    c=sqlite3.connect(DB,check_same_thread=False)
    c.execute("""CREATE TABLE IF NOT EXISTS projects(id TEXT PRIMARY KEY,title TEXT,focus TEXT,created TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS interviews(id TEXT PRIMARY KEY,project_id TEXT,informant TEXT,filename TEXT,question TEXT,quote TEXT,created TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS coding(id TEXT PRIMARY KEY,project_id TEXT,data_id TEXT,open_code TEXT,category TEXT,theme TEXT,memo TEXT,status TEXT,updated TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS observations(id TEXT PRIMARY KEY,project_id TEXT,date TEXT,location TEXT,focus TEXT,note TEXT,theme TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS documents(id TEXT PRIMARY KEY,project_id TEXT,filename TEXT,path TEXT,date TEXT,location TEXT,description TEXT,theme TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS triangulation(id TEXT PRIMARY KEY,project_id TEXT,theme TEXT,interview TEXT,observation TEXT,documentation TEXT,interpretation TEXT,status TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS audit(id TEXT PRIMARY KEY,project_id TEXT,action TEXT,detail TEXT,created TEXT)""")
    c.commit(); return c

db=conn()
def audit(action,detail):
    db.execute("INSERT INTO audit VALUES(?,?,?,?,?)",(str(uuid.uuid4()),st.session_state.project_id,action,detail,datetime.now().isoformat(timespec="seconds")))
    db.commit()
def qdf(sql,args=()):
    return pd.read_sql_query(sql,db,params=args)
def save_project(title,focus):
    pid=st.session_state.project_id
    db.execute("INSERT OR REPLACE INTO projects VALUES(?,?,?,?)",(pid,title,focus,datetime.now().isoformat(timespec="seconds")));db.commit()
def records():
    return qdf("SELECT * FROM interviews WHERE project_id=? ORDER BY created",(st.session_state.project_id,))
def coding_df():
    return qdf("""SELECT i.id AS data_id,i.informant AS 'Kode informan',i.filename AS 'Nama file',i.question AS Pertanyaan,i.quote AS 'Kutipan asli',
    COALESCE(c.open_code,'') AS 'Open code',COALESCE(c.category,'') AS Kategori,COALESCE(c.theme,'') AS Tema,
    COALESCE(c.memo,'') AS Memo,COALESCE(c.status,'Belum ditinjau') AS Status
    FROM interviews i LEFT JOIN coding c ON c.data_id=i.id WHERE i.project_id=? ORDER BY i.created""",(st.session_state.project_id,))
def upsert_coding(row):
    db.execute("""INSERT OR REPLACE INTO coding VALUES(?,?,?,?,?,?,?,?,?)""",
      (str(uuid.uuid4()),st.session_state.project_id,row["data_id"],str(row["Open code"]),str(row["Kategori"]),str(row["Tema"]),str(row["Memo"]),str(row["Status"]),datetime.now().isoformat(timespec="seconds")))
    db.commit()
def clean(v): return "" if pd.isna(v) else str(v).strip()
def find_col(cols,keys,default):
    for c in cols:
        if any(k in re.sub(r'[^a-z0-9]','',str(c).lower()) for k in keys): return c
    return cols[min(default,len(cols)-1)]

if "project_id" not in st.session_state: st.session_state.project_id=str(uuid.uuid4())
projects=qdf("SELECT * FROM projects ORDER BY created DESC")
st.title("Q-ANALYST")
st.caption("Personal Qualitative Research Analyzer · Local-first · Multi-file Excel · AI-assisted")
with st.sidebar:
    st.header("Ruang penelitian")
    if len(projects):
        opts={f"{r.title} · {r.id[:6]}":r.id for r in projects.itertuples()}
        chosen=st.selectbox("Pilih proyek",list(opts.keys()))
        if st.button("Buka proyek terpilih"): st.session_state.project_id=opts[chosen]; st.rerun()
    st.caption("Database SQLite tersimpan lokal di folder aplikasi.")
    st.divider()
    st.caption("AI eksternal bersifat opsional. Data tidak dikirim ke AI kecuali Bapak menjalankan fitur tersebut.")

tabs=st.tabs(["Dashboard & Proyek","Impor Multi-Excel","AI Coding","Observasi","Dokumentasi","Matriks Lintas Informan","Triangulasi","Laporan & Audit"])

with tabs[0]:
    st.subheader("Identitas proyek")
    existing=qdf("SELECT * FROM projects WHERE id=?",(st.session_state.project_id,))
    title=st.text_input("Judul penelitian",existing.iloc[0]["title"] if len(existing) else "Penelitian Studi Kasus")
    focus=st.text_area("Fokus / pertanyaan penelitian",existing.iloc[0]["focus"] if len(existing) else "")
    if st.button("Simpan proyek",type="primary"):
        save_project(title,focus);audit("Simpan proyek",title);st.success("Proyek tersimpan di database lokal.")
    data=records(); codes=coding_df()
    obs=qdf("SELECT * FROM observations WHERE project_id=?",(st.session_state.project_id,))
    docs=qdf("SELECT * FROM documents WHERE project_id=?",(st.session_state.project_id,))
    a,b,c,d=st.columns(4)
    a.metric("Unit wawancara",len(data));b.metric("Unit coding",len(codes));c.metric("Catatan observasi",len(obs));d.metric("Dokumentasi",len(docs))
    st.info("Alur: proyek → impor beberapa Excel → validasi data → AI coding opsional → observasi/dokumentasi → perbandingan → triangulasi → laporan.")
    st.warning("Gunakan kode informan (I-01, I-02, dst.). Hindari menyimpan nama asli, nomor telepon, alamat, atau informasi sensitif yang tidak diperlukan.")

with tabs[1]:
    st.subheader("Impor beberapa file Excel sekaligus")
    files=st.file_uploader("Pilih satu atau banyak file Excel",type=["xlsx","xls"],accept_multiple_files=True)
    if files:
        allrows=[]
        for f in files:
            try:
                xl=pd.ExcelFile(f)
                sheet=st.selectbox(f"Sheet — {f.name}",xl.sheet_names,key="sh"+hashlib.md5(f.name.encode()).hexdigest())
                d=pd.read_excel(f,sheet_name=sheet).fillna("")
                st.write(f"**{f.name}** · {len(d)} baris");st.dataframe(d.head(5),use_container_width=True)
                cols=list(d.columns)
                ic=st.selectbox(f"Kolom informan — {f.name}",cols,index=cols.index(find_col(cols,["informan","responden","narasumber","kode"],0)),key="i"+f.name)
                qc=st.selectbox(f"Kolom pertanyaan — {f.name}",cols,index=cols.index(find_col(cols,["pertanyaan","question"],min(1,len(cols)-1))),key="q"+f.name)
                ac=st.selectbox(f"Kolom jawaban — {f.name}",cols,index=cols.index(find_col(cols,["jawaban","answer","respon"],min(2,len(cols)-1))),key="a"+f.name)
                for _,r in d.iterrows():
                    if clean(r[ac]):
                        allrows.append((str(uuid.uuid4()),st.session_state.project_id,clean(r[ic]),f.name,clean(r[qc]),clean(r[ac]),datetime.now().isoformat(timespec="seconds")))
            except Exception as e: st.error(f"Gagal membaca {f.name}: {e}")
        if st.button("Simpan dan gabungkan file terpilih",type="primary"):
            if allrows:
                db.executemany("INSERT INTO interviews VALUES(?,?,?,?,?,?,?)",allrows);db.commit()
                audit("Impor wawancara",f"{len(allrows)} unit dari {len(files)} file")
                st.success(f"{len(allrows)} unit jawaban tersimpan permanen di database lokal.")
    st.dataframe(records()[["informant","filename","question","quote"]] if len(records()) else pd.DataFrame(),use_container_width=True)

with tabs[2]:
    st.subheader("AI-assisted coding")
    st.write("AI mengusulkan open code, kategori, tema, dan memo. Semua usulan harus divalidasi peneliti.")
    api=st.text_input("OpenAI API key (opsional; tidak disimpan ke database)",type="password")
    model=st.text_input("Nama model API",value="gpt-4.1-mini")
    focus_current=existing.iloc[0]["focus"] if len(existing) else ""
    batch=st.slider("Unit per batch",1,15,6)
    data=coding_df()
    if not len(data): st.info("Impor data wawancara terlebih dahulu.")
    elif st.button("Jalankan AI untuk unit yang belum memiliki coding",type="primary",disabled=not bool(api)):
        try:
            from openai import OpenAI
            client=OpenAI(api_key=api)
            todo=data[data["Open code"].astype(str).str.strip()==""]
            bar=st.progress(0);msg=st.empty()
            for start in range(0,len(todo),batch):
                part=todo.iloc[start:start+batch]
                payload=[{"id":r.data_id,"informan":r["Kode informan"],"pertanyaan":r.Pertanyaan,"kutipan":r["Kutipan asli"]} for _,r in part.iterrows()]
                prompt=f"""Anda asisten metodologi penelitian kualitatif studi kasus. Fokus: {focus_current or 'induktif, belum ditentukan'}.
Analisis hanya bukti yang diberikan. Jangan membuat data/kutipan/konteks. Untuk tiap ID berikan unit_makna, open_code, kategori, tema, memo, confidence, dan alasan.
Gunakan bahasa Indonesia akademik, kode dekat dengan data, jangan memaksakan teori. Jika bukti ambigu, nyatakan perlu klarifikasi. Output JSON murni: {{ "items":[{{"id":"...","open_code":"...","kategori":"...","tema":"...","memo":"...","confidence":"...","alasan":"..."}}] }}.
DATA: {json.dumps(payload,ensure_ascii=False)}"""
                resp=client.responses.create(model=model,input=prompt)
                txt=resp.output_text
                obj=json.loads(txt[txt.find("{"):txt.rfind("}")+1])
                for it in obj.get("items",[]):
                    ix=data[data.data_id==it.get("id")]
                    if len(ix):
                        r=ix.iloc[0]
                        upsert_coding({"data_id":r.data_id,"Open code":it.get("open_code",""),"Kategori":it.get("kategori",""),"Tema":it.get("tema",""),"Memo":f"{it.get('memo','')} | Keyakinan: {it.get('confidence','')}; alasan: {it.get('alasan','')}","Status":"Usulan AI — perlu validasi"})
                bar.progress(min(1,(start+batch)/max(1,len(todo))));msg.write(f"Memproses {min(start+batch,len(todo))}/{len(todo)}")
            audit("AI coding",f"{len(todo)} unit diproses")
            st.success("Usulan coding selesai. Periksa dan validasi setiap baris.")
        except Exception as e: st.error(f"AI gagal: {e}")
    data=coding_df()
    if len(data):
        edited=st.data_editor(data,use_container_width=True,hide_index=True,key="codingedit")
        if st.button("Simpan hasil validasi coding"):
            for _,r in edited.iterrows(): upsert_coding(r)
            audit("Validasi coding","Perubahan coding disimpan")
            st.success("Coding tersimpan.")
    st.caption("Teks kutipan yang diproses akan dikirim ke layanan AI eksternal. Samarkan identitas sebelum menggunakannya. Biaya API mengikuti akun penyedia.")

with tabs[3]:
    st.subheader("Catatan observasi")
    with st.form("obs",clear_on_submit=True):
        dt=st.date_input("Tanggal",date.today());loc=st.text_input("Lokasi/konteks");foc=st.text_input("Fokus observasi")
        note=st.text_area("Catatan lapangan");theme=st.text_input("Tema/kode terkait")
        ok=st.form_submit_button("Simpan observasi")
    if ok:
        db.execute("INSERT INTO observations VALUES(?,?,?,?,?,?,?)",(str(uuid.uuid4()),st.session_state.project_id,str(dt),loc,foc,note,theme));db.commit();audit("Tambah observasi",f"{dt} {foc}");st.success("Catatan observasi tersimpan.")
    st.dataframe(obs if 'obs' in locals() else qdf("SELECT * FROM observations WHERE project_id=?",(st.session_state.project_id,)),use_container_width=True)

with tabs[4]:
    st.subheader("Dokumentasi foto dan deskripsi")
    imgs=st.file_uploader("Unggah beberapa foto JPG/PNG",type=["jpg","jpeg","png"],accept_multiple_files=True)
    if imgs:
        for f in imgs:
            raw=f.getvalue();fid=str(uuid.uuid4());path=UPLOAD/f"{fid}_{Path(f.name).name}"
            path.write_bytes(raw)
            db.execute("INSERT INTO documents VALUES(?,?,?,?,?,?,?,?)",(fid,st.session_state.project_id,f.name,str(path),str(date.today()),"","",""));db.commit()
        audit("Unggah dokumentasi",f"{len(imgs)} foto")
        st.success("Foto disimpan di folder uploads lokal.")
    docs=qdf("SELECT * FROM documents WHERE project_id=?",(st.session_state.project_id,))
    for _,r in docs.iterrows():
        with st.expander(r["filename"]):
            p=Path(r["path"])
            if p.exists(): st.image(str(p),width=250)
            desc=st.text_area("Deskripsi/konteks",r["description"],key="desc"+r["id"])
            loc=st.text_input("Lokasi",r["location"],key="loc"+r["id"])
            th=st.text_input("Tema terkait",r["theme"],key="th"+r["id"])
            if st.button("Simpan metadata foto",key="save"+r["id"]):
                db.execute("UPDATE documents SET description=?,location=?,theme=? WHERE id=?",(desc,loc,th,r["id"]));db.commit();audit("Edit metadata foto",r["filename"]);st.success("Metadata tersimpan.")
    st.caption("Foto disimpan lokal. AI visual/OCR belum aktif; interpretasi foto dilakukan peneliti berdasarkan konteks lapangan.")

with tabs[5]:
    st.subheader("Matriks lintas informan")
    cd=coding_df()
    if len(cd) and cd["Tema"].astype(str).str.strip().ne("").any():
        pivot=pd.pivot_table(cd[cd["Tema"].astype(str).str.strip()!=""],index="Tema",columns="Kode informan",values="data_id",aggfunc="count",fill_value=0)
        st.dataframe(pivot,use_container_width=True)
        st.caption("Angka adalah jumlah unit data per tema dan informan, bukan skor persetujuan atau bobot tema.")
    else: st.info("Tema belum tersedia. Lengkapi coding terlebih dahulu.")

with tabs[6]:
    st.subheader("Matriks triangulasi")
    with st.form("tri",clear_on_submit=True):
        th=st.text_input("Tema/fokus yang ditriangulasi")
        wi=st.text_area("Bukti wawancara")
        ob=st.text_area("Bukti observasi")
        dc=st.text_area("Bukti dokumentasi")
        inter=st.text_area("Interpretasi: konvergensi, perbedaan, kontradiksi, kekosongan bukti")
        status=st.selectbox("Status pemeriksaan",["Belum diperiksa","Bukti saling menguatkan","Bukti berbeda/perlu penjelasan","Bukti belum memadai"])
        ok=st.form_submit_button("Simpan matriks triangulasi")
    if ok:
        db.execute("INSERT INTO triangulation VALUES(?,?,?,?,?,?,?,?)",(str(uuid.uuid4()),st.session_state.project_id,th,wi,ob,dc,inter,status));db.commit();audit("Simpan triangulasi",th);st.success("Matriks triangulasi tersimpan.")
    st.dataframe(qdf("SELECT theme,interview,observation,documentation,interpretation,status FROM triangulation WHERE project_id=?",(st.session_state.project_id,)),use_container_width=True)
    st.caption("Triangulasi bukan sekadar menemukan kesamaan; perbedaan dan bukti negatif juga harus dianalisis.")

with tabs[7]:
    st.subheader("Ekspor dan jejak audit")
    if st.button("Siapkan workbook Excel"):
        b=BytesIO()
        with pd.ExcelWriter(b,engine="openpyxl") as w:
            records().to_excel(w,index=False,sheet_name="Wawancara")
            coding_df().to_excel(w,index=False,sheet_name="Coding")
            qdf("SELECT * FROM observations WHERE project_id=?",(st.session_state.project_id,)).to_excel(w,index=False,sheet_name="Observasi")
            qdf("SELECT filename,date,location,description,theme FROM documents WHERE project_id=?",(st.session_state.project_id,)).to_excel(w,index=False,sheet_name="Dokumentasi")
            qdf("SELECT theme,interview,observation,documentation,interpretation,status FROM triangulation WHERE project_id=?",(st.session_state.project_id,)).to_excel(w,index=False,sheet_name="Triangulasi")
        st.download_button("Unduh hasil Excel",b.getvalue(),"Q_ANALYST_Hasil.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if st.button("Siapkan laporan Word"):
        doc=Document();doc.add_heading("Laporan Kerja Analisis Data Kualitatif",0)
        doc.add_paragraph(title if 'title' in locals() else st.session_state.project_id)
        doc.add_paragraph(f"Jumlah unit wawancara: {len(records())}")
        doc.add_heading("Matriks coding",level=1)
        cd=coding_df()
        if len(cd):
            t=doc.add_table(rows=1,cols=5);t.style="Table Grid"
            for c,v in zip(t.rows[0].cells,["Informan","Kutipan asli","Open code","Kategori","Tema"]):c.text=v
            for _,r in cd.iterrows():
                cells=t.add_row().cells
                for c,v in zip(cells,[r["Kode informan"],r["Kutipan asli"],r["Open code"],r["Kategori"],r["Tema"]]):c.text=str(v)[:1000]
        b=BytesIO();doc.save(b);st.download_button("Unduh laporan Word",b.getvalue(),"Q_ANALYST_Laporan.docx","application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    st.markdown("**Audit trail**")
    st.dataframe(qdf("SELECT action,detail,created FROM audit WHERE project_id=? ORDER BY created DESC",(st.session_state.project_id,)),use_container_width=True)
    st.caption("Database SQLite dan foto berada di folder aplikasi. Backup folder aplikasi secara berkala. Ekspor laporan ini adalah draf kerja, bukan naskah Bab IV final otomatis.")
