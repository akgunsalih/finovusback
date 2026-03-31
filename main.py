from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
import openpyxl
from io import BytesIO
from datetime import datetime, date
import json
from typing import List, Optional, Dict
from pydantic import BaseModel
import asyncio
from fastapi import WebSocket, WebSocketDisconnect
import random

app = FastAPI(title="Finovus API")

# Enable CORS for frontend development
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- WebSocket & Global State ---
class ConnectionManager:
    def __init__(self):
        self.active_connections: List[WebSocket] = []

    async def lock_connect(self, websocket: WebSocket):
        await websocket.accept()
        self.active_connections.append(websocket)

    def disconnect(self, websocket: WebSocket):
        self.active_connections.remove(websocket)

    async def broadcast(self, message: dict):
        for connection in self.active_connections:
            try:
                await connection.send_json(message)
            except:
                continue

manager = ConnectionManager()
latest_data = {"meta": None, "sonuclar": []}
simulation_task = None

# --- Models ---
class SonucRow(BaseModel):
    kontrat: Optional[str]
    aciklama: Optional[str]
    alis: Optional[float]
    spot_satis: Optional[float]
    gun: Optional[int]
    hesaplama: Optional[float]
    referans_faiz: Optional[float]
    islem_onerisi: Optional[str]

class MetaInfo(BaseModel):
    bugun: Optional[str]
    referans_faiz: Optional[float]
    toplam_satir: int
    islem_yap: int
    islem_yapma: int

class CalculationResult(BaseModel):
    meta: MetaInfo
    sonuclar: List[SonucRow]

# --- Helper Functions (Migrated from finovus_hesapla.py) ---
def to_date(val) -> Optional[date]:
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    return None

def sheet_to_list(ws) -> List[Dict]:
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    result = []
    for row in rows[1:]:
        result.append(dict(zip(headers, row)))
    return result

def perform_calculation(file_content: bytes) -> dict:
    wb = openpyxl.load_workbook(BytesIO(file_content), data_only=True, keep_links=False)
    
    # REFERANS FAİZ
    if "REFERANS FAİZ" not in wb.sheetnames:
        raise ValueError("REFERANS FAİZ sayfası bulunamadı.")
    
    rf_rows = sheet_to_list(wb["REFERANS FAİZ"])
    referans_faiz = None
    for r in rf_rows:
        if str(r.get("KOD", "")).strip() == "TLREF":
            try:
                referans_faiz = float(r["FAİZ"])
            except:
                pass
            break
            
    # SÖZLEŞME TARİH
    if "SÖZLEŞME TARİH" not in wb.sheetnames:
        raise ValueError("SÖZLEŞME TARİH sayfası bulunamadı.")
    
    st_rows = sheet_to_list(wb["SÖZLEŞME TARİH"])
    sozlesme = {}
    for r in st_rows:
        key = str(r.get("TARİH", "")).strip()
        val = to_date(r.get("VADE SONU"))
        if key and val:
            sozlesme[key] = val
    
    bugun_tarihi = sozlesme.get("Bugün")
    
    # MATRİKS VERİ SPOT
    if "MATRİKS VERİ SPOT" not in wb.sheetnames:
        raise ValueError("MATRİKS VERİ SPOT sayfası bulunamadı.")
        
    spot_rows = sheet_to_list(wb["MATRİKS VERİ SPOT"])
    spot = {}
    for r in spot_rows:
        sembol = str(r.get("SEMBOL", "")).strip().upper()
        try:
            spot[sembol] = float(r["SATIŞ"])
        except:
            continue
            
    # MATRİKS VERİ VADELİ
    if "MATRİKS VERİ VADELİ" not in wb.sheetnames:
        raise ValueError("MATRİKS VERİ VADELİ sayfası bulunamadı.")
        
    vadeli_rows = sheet_to_list(wb["MATRİKS VERİ VADELİ"])
    sonuclar = []
    
    for row in vadeli_rows:
        kontrat = str(row.get("KONTRAT", "") or "").strip()
        aciklama = str(row.get("AÇIKLAMA", "") or "").strip()
        
        try:
            alis = float(row.get("ALIŞ") or 0)
        except:
            alis = None
            
        kelimeler = aciklama.split()
        ilk_kelime = kelimeler[0].upper() if len(kelimeler) >= 1 else ""
        ikinci_kel = kelimeler[1].strip() if len(kelimeler) >= 2 else ""
        
        spot_satis = spot.get(ilk_kelime) if ilk_kelime else None
        
        gun = None
        if ikinci_kel and bugun_tarihi:
            vade_sonu = sozlesme.get(ikinci_kel)
            if vade_sonu:
                delta = (vade_sonu - bugun_tarihi).days
                gun = delta if delta > 0 else None
                
        hesaplama = None
        if alis and spot_satis and spot_satis != 0 and gun and gun != 0:
            hesaplama = ((spot_satis / alis) - 1) / gun * 365
            
        islem_onerisi = None
        if hesaplama is not None and referans_faiz is not None:
            hesaplama_pct = hesaplama * 100
            islem_onerisi = "İŞLEM YAP" if hesaplama_pct > referans_faiz else "İŞLEM YAPMA"
            
        sonuclar.append({
            "kontrat": kontrat,
            "aciklama": aciklama,
            "alis": alis,
            "spot_satis": spot_satis,
            "gun": gun,
            "hesaplama": round(hesaplama * 100, 4) if hesaplama is not None else None,
            "referans_faiz": referans_faiz,
            "islem_onerisi": islem_onerisi
        })
        
    islem_yap = sum(1 for r in sonuclar if r["islem_onerisi"] == "İŞLEM YAP")
    islem_yapma = sum(1 for r in sonuclar if r["islem_onerisi"] == "İŞLEM YAPMA")
    
    return {
        "meta": {
            "bugun": str(bugun_tarihi) if bugun_tarihi else None,
            "referans_faiz": referans_faiz,
            "toplam_satir": len(sonuclar),
            "islem_yap": islem_yap,
            "islem_yapma": islem_yapma
        },
        "sonuclar": sonuclar
    }

async def run_simulation():
    global latest_data
    while True:
        await asyncio.sleep(3) # Her 3 saniyede bir güncelle
        if not latest_data["sonuclar"]:
            continue
            
        # Bazı rakamları hafifçe oynat (Spot Satış ve Alış %0.1 civarı)
        for row in latest_data["sonuclar"]:
            if row["alis"]:
                row["alis"] *= (1 + (random.random() - 0.5) * 0.001)
            if row["spot_satis"]:
                row["spot_satis"] *= (1 + (random.random() - 0.5) * 0.001)
            
            # Yeniden hesapla
            if row["alis"] and row["spot_satis"] and row["gun"]:
                hesaplama = ((row["spot_satis"] / row["alis"]) - 1) / row["gun"] * 365
                row["hesaplama"] = round(hesaplama * 100, 4)
                
                ref = latest_data["meta"]["referans_faiz"]
                if ref is not None:
                    row["islem_onerisi"] = "İŞLEM YAP" if row["hesaplama"] > ref else "İŞLEM YAPMA"

        # Metaları güncelle
        latest_data["meta"]["islem_yap"] = sum(1 for r in latest_data["sonuclar"] if r["islem_onerisi"] == "İŞLEM YAP")
        latest_data["meta"]["islem_yapma"] = sum(1 for r in latest_data["sonuclar"] if r["islem_onerisi"] == "İŞLEM YAPMA")
        
        await manager.broadcast(latest_data)

def generate_excel(result: dict) -> BytesIO:
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb_out = openpyxl.Workbook()
    ws = wb_out.active
    ws.title = "SONUÇLAR"

    headers = ["KONTRAT", "AÇIKLAMA", "ALIŞ", "SPOT SATIŞ", "GÜN",
               "HESAPLAMA %", "REFERANS FAİZ %", "İŞLEM ÖNERİSİ"]
    ws.append(headers)

    hdr_fill = PatternFill("solid", fgColor="1F3864")
    hdr_font = Font(bold=True, color="FFFFFF", size=11)
    for cell in ws[1]:
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    yesil = PatternFill("solid", fgColor="C6EFCE")
    kirmizi = PatternFill("solid", fgColor="FFC7CE")
    gri = PatternFill("solid", fgColor="F2F2F2")
    beyaz = PatternFill("solid", fgColor="FFFFFF")
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for i, r in enumerate(result["sonuclar"], start=2):
        ws.append([
            r["kontrat"], r["aciklama"], r["alis"], r["spot_satis"],
            r["gun"], r["hesaplama"], r["referans_faiz"], r["islem_onerisi"],
        ])
        fill = gri if i % 2 == 0 else beyaz
        for cell in ws[i]:
            cell.fill = fill
            cell.alignment = Alignment(vertical="center")
            cell.border = border

        onerisi_cell = ws.cell(row=i, column=8)
        if r["islem_onerisi"] == "İŞLEM YAP":
            onerisi_cell.fill = yesil
            onerisi_cell.font = Font(bold=True, color="276221")
        elif r["islem_onerisi"] == "İŞLEM YAPMA":
            onerisi_cell.fill = kirmizi
            onerisi_cell.font = Font(bold=True, color="9C0006")

        for col in [3, 4, 6, 7]:
            c = ws.cell(row=i, column=col)
            if c.value is not None:
                c.number_format = '0.0000' if col >= 6 else '0.00'

    for hdr_cell in ws[1]:
        hdr_cell.border = border

    col_widths = [18, 38, 10, 12, 7, 14, 16, 18]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 25
    ws.freeze_panes = "A2"

    output = BytesIO()
    wb_out.save(output)
    output.seek(0)
    return output

# --- Endpoints ---
@app.get("/")
async def root():
    return {"status": "ok", "message": "Finovus API is running"}

@app.post("/calculate", response_model=CalculationResult)
async def calculate(file: UploadFile = File(...)):
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Lütfen geçerli bir Excel dosyası yükleyin.")
    
    try:
        content = await file.read()
        global latest_data, simulation_task
        latest_data = perform_calculation(content)
        
        # Simülasyonu başlat (eğer başlamadıysa)
        if simulation_task is None:
            simulation_task = asyncio.create_task(run_simulation())
            
        return latest_data
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/export")
async def export(file: UploadFile = File(...)):
    try:
        content = await file.read()
        result = perform_calculation(content)
        excel_file = generate_excel(result)
        
        filename = f"FINOVUS_SONUC_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return StreamingResponse(
            excel_file,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket):
    await manager.lock_connect(websocket)
    try:
        # Bağlanınca varsa son veriyi hemen gönder
        if latest_data["sonuclar"]:
            await websocket.send_json(latest_data)
        while True:
            await websocket.receive_text() # Bağlantıyı canlı tutmak için bekler
    except WebSocketDisconnect:
        manager.disconnect(websocket)

class ExportData(BaseModel):
    meta: MetaInfo
    sonuclar: List[SonucRow]

@app.post("/export-json")
async def export_json(data: ExportData):
    try:
        excel_file = generate_excel(data.dict())
        filename = f"FINOVUS_CANLI_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return StreamingResponse(
            excel_file,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
