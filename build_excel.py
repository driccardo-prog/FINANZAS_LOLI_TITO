# -*- coding: utf-8 -*-
"""
Genera FINANZAS_LOLI_TITO.xlsx
Hoja 1: INGRESOS (editable)
Hojas de meses: JUNIO..DICIEMBRE (arranca con data real de Jun/Jul)
Hoja CONFIG: mapa de categorias para auto-clasificar el resumen AMEX
Hoja RESUMEN ANUAL
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment

# ---------- Paleta ----------
AZUL      = "1F3864"   # titulos
AZUL2     = "2E4C7E"
GRIS_HDR  = "44546A"   # headers de tabla
GRIS_CL   = "D9D9D9"   # gris claro (sugerido)
GRIS_MED  = "BFBFBF"
AMAR      = "FFF2CC"   # editable
AMAR_BORDE= "FFD966"
VERDE     = "C6EFCE"   # sobro
VERDE_TXT = "006100"
ROJO      = "FFC7CE"   # falto
ROJO_TXT  = "9C0006"
BLANCO    = "FFFFFF"
CELESTE   = "DDEBF7"   # subtotales / secciones
CELESTE2  = "BDD7EE"

thin = Side(style="thin", color="BFBFBF")
med  = Side(style="medium", color="808080")
border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
border_box = Border(left=med, right=med, top=med, bottom=med)

def fill(c): return PatternFill("solid", fgColor=c)

def style(cell, *, bold=False, size=11, color="000000", bg=None, align="left",
          border=None, italic=False, wrap=False, numfmt=None):
    cell.font = Font(bold=bold, size=size, color=color, italic=italic, name="Calibri")
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if bg: cell.fill = fill(bg)
    if border: cell.border = border
    if numfmt: cell.number_format = numfmt
    return cell

MONEY = '#,##0'
MONEYAR = '"$"#,##0'
PCT = '0.0%'

wb = Workbook()

# ======================================================================
# HOJA CONFIG (mapa de auto-categorizacion AMEX)
# ======================================================================
cfg = wb.active
cfg.title = "CONFIG"
cfg["A1"] = "MAPA DE CATEGORÍAS · Resumen AMEX (auto)"
cfg.merge_cells("A1:B1")
style(cfg["A1"], bold=True, size=13, color=BLANCO, bg=AZUL, align="center")
cfg["A2"] = "Palabra clave (aparece en el detalle)"
cfg["B2"] = "Categoría"
for c in ("A2","B2"): style(cfg[c], bold=True, color=BLANCO, bg=GRIS_HDR, align="center", border=border_all)

# palabra clave -> categoria (se busca si el texto del comercio CONTIENE la palabra)
mapa = [
    ("rappi", "Delivery"), ("pedidos ya", "Delivery"), ("pedidosya", "Delivery"),
    ("foodgolf", "Delivery"), ("dlorappi", "Delivery"), ("uber eats", "Delivery"),
    ("delfin", "Supermercado"), ("manzana", "Supermercado"), ("manzanares", "Supermercado"),
    ("vacalin", "Supermercado"), ("paroissien", "Supermercado"), ("demente", "Supermercado"),
    ("quesos", "Supermercado"), ("coto", "Supermercado"), ("carrefour", "Supermercado"),
    ("dia", "Supermercado"), ("jumbo", "Supermercado"), ("disco", "Supermercado"),
    ("chango", "Supermercado"), ("vea", "Supermercado"), ("market", "Supermercado"),
    ("rosetta", "Supermercado"),
    ("verdu", "Verdulería"), ("fruteria", "Verdulería"), ("fruta", "Verdulería"),
    ("café", "Cafecito"), ("cafe", "Cafecito"), ("solileb", "Cafecito"),
    ("noire", "Cafecito"), ("razo", "Cafecito"), ("entre nos", "Cafecito"),
    ("starbucks", "Cafecito"), ("havanna", "Cafecito"),
    ("parmegiano", "Restaurant"), ("santicheese", "Restaurant"), ("evelia", "Restaurant"),
    ("hamburgues", "Restaurant"), ("garcia", "Restaurant"), ("resto", "Restaurant"),
    ("parrilla", "Restaurant"), ("pizza", "Restaurant"), ("sushi", "Restaurant"),
    ("groovypets", "Veterinaria"), ("veterin", "Veterinaria"), ("mascota", "Veterinaria"),
    ("royal canin", "Veterinaria"), ("pet", "Veterinaria"),
    ("ypf", "Nafta"), ("shell", "Nafta"), ("axion", "Nafta"), ("puma", "Nafta"), ("nafta", "Nafta"),
    ("jetsmart", "Viaje"), ("latam", "Viaje"), ("aerolineas", "Viaje"),
    ("despegar", "Viaje"), ("almundo", "Viaje"), ("booking", "Viaje"), ("airbnb", "Viaje"),
    ("farmacia", "Farmacia"), ("soytufarmacia", "Farmacia"), ("farmacity", "Farmacia"),
    ("cabify", "Transporte"), ("uber", "Transporte"), ("sube", "Transporte"), ("didi", "Transporte"),
    ("rentas", "Trámites"), ("afip", "Trámites"), ("registro", "Trámites"), ("tramite", "Trámites"),
    ("h2soap", "Otros"), ("monzon", "Otros"),
]
r = 3
for kw, cat in mapa:
    cfg.cell(r,1,kw); cfg.cell(r,2,cat)
    style(cfg.cell(r,1), border=border_all)
    style(cfg.cell(r,2), border=border_all)
    r += 1
MAP_LAST = r + 20  # dejar filas libres para agregar
for rr in range(r, MAP_LAST):
    style(cfg.cell(rr,1), border=border_all, bg=AMAR)
    style(cfg.cell(rr,2), border=border_all, bg=AMAR)
cfg.column_dimensions["A"].width = 34
cfg.column_dimensions["B"].width = 20
cfg["H2"] = "Agregá acá abajo (en amarillo) cualquier comercio nuevo y su categoría. Los meses lo usan para acomodar solo el resumen de la AMEX."
cfg.merge_cells("H2:K6")
style(cfg["H2"], italic=True, wrap=True, color="595959", align="left")
cfg.column_dimensions["H"].width = 20
MAP_RANGE_KW  = f"CONFIG!$A$3:$A${MAP_LAST-1}"
MAP_RANGE_CAT = f"CONFIG!$B$3:$B${MAP_LAST-1}"

# lista de categorias para validaciones/resumen
CATS = ["Supermercado","Verdulería","Delivery","Cafecito","Restaurant","Salida",
        "Nafta","Transporte","Veterinaria","Farmacia","Viaje","Trámites","Otros"]
cfg["A1"].comment = Comment("Esta hoja hace que el resumen de la AMEX se acomode solo por categoría.", "Finanzas")
# guardo lista de categorias en columna F para validacion
cfg["F1"] = "CATEGORÍAS"
style(cfg["F1"], bold=True, color=BLANCO, bg=GRIS_HDR, align="center")
for i,ct in enumerate(CATS):
    cfg.cell(2+i,6,ct); style(cfg.cell(2+i,6), border=border_all)
CAT_LIST_RANGE = f"CONFIG!$F$2:$F${1+len(CATS)}"
cfg.column_dimensions["F"].width = 16

# ======================================================================
# HOJA 1: INGRESOS
# ======================================================================
ing = wb.create_sheet("INGRESOS")
for col,w in zip("ABCDE",[26,16,12,30,18]): ing.column_dimensions[col].width = w

ing.merge_cells("A1:E1")
ing["A1"] = "FINANZAS · LOLI & TITO"
style(ing["A1"], bold=True, size=18, color=BLANCO, bg=AZUL, align="center")
ing.row_dimensions[1].height = 30

ing.merge_cells("A3:E3")
ing["A3"] = "INGRESOS NETOS MENSUALES"
style(ing["A3"], bold=True, size=12, color=BLANCO, bg=AZUL2, align="center")

hdr = ["Persona","Ingreso","% real"]
for i,h in enumerate(hdr):
    style(ing.cell(4,1+i,h), bold=True, color=BLANCO, bg=GRIS_HDR, align="center", border=border_all)
ing.merge_cells("D4:E4"); style(ing["D4"], bold=True, color=BLANCO, bg=GRIS_HDR, align="center", border=border_all)
ing["D4"]="Proporcional real (referencia)"

ing["A5"]="Loli"; ing["B5"]=2800000; ing["C5"]="=B5/$B$7"
ing["A6"]="Tito"; ing["B6"]=4000000; ing["C6"]="=B6/$B$7"
ing["A7"]="Total"; ing["B7"]="=B5+B6"; ing["C7"]="=B7/$B$7"
for rr in (5,6):
    style(ing.cell(rr,1), border=border_all)
    style(ing.cell(rr,2), border=border_all, bg=AMAR, numfmt=MONEYAR, align="right")
    style(ing.cell(rr,3), border=border_all, numfmt=PCT, align="center")
    ing.merge_cells(f"D{rr}:E{rr}"); style(ing.cell(rr,4), border=border_all)
style(ing["A7"], bold=True, border=border_all, bg=CELESTE)
style(ing["B7"], bold=True, border=border_all, bg=CELESTE, numfmt=MONEYAR, align="right")
style(ing["C7"], bold=True, border=border_all, bg=CELESTE, numfmt=PCT, align="center")
ing.merge_cells("D7:E7"); style(ing["D7"], border=border_all, bg=CELESTE)

# SPLIT EN USO
ing.merge_cells("A9:E9")
ing["A9"]="SPLIT EN USO (editable)"
style(ing["A9"], bold=True, size=12, color=BLANCO, bg=AZUL2, align="center")
for i,h in enumerate(["Concepto","% Loli","% Tito","Notas"]):
    style(ing.cell(10,1+i,h), bold=True, color=BLANCO, bg=GRIS_HDR, align="center", border=border_all)
ing.merge_cells("D10:E10")
ing["A11"]="Split actual"; ing["B11"]=0.40; ing["C11"]="=1-B11"
ing.merge_cells("D11:E11")
ing["D11"]="Cambiá solo la celda amarilla (% Loli). Tito se ajusta solo."
style(ing["A11"], border=border_all)
style(ing["B11"], border=border_all, bg=AMAR, numfmt=PCT, align="center", bold=True)
style(ing["C11"], border=border_all, numfmt=PCT, align="center", bold=True)
style(ing["D11"], border=border_all, italic=True, wrap=True, color="595959")
ing.row_dimensions[11].height = 28
# nombre de celda para split
wb.defined_names.add.__self__  # no-op guard
from openpyxl.workbook.defined_name import DefinedName
wb.defined_names.add(DefinedName("PCT_LOLI", attr_text="INGRESOS!$B$11"))
wb.defined_names.add(DefinedName("PCT_TITO", attr_text="INGRESOS!$C$11"))

# MONTOS FIJOS DE REFERENCIA
ing.merge_cells("A13:E13")
ing["A13"]="MONTOS FIJOS DE REFERENCIA"
style(ing["A13"], bold=True, size=12, color=BLANCO, bg=AZUL2, align="center")
for i,h in enumerate(["Concepto","Núñez","Pilar"]):
    style(ing.cell(14,1+i,h), bold=True, color=BLANCO, bg=GRIS_HDR, align="center", border=border_all)
ref = [
    ("Expensas",415277,336690),("Luz (EDENOR)",52392,None),("Gas",None,None),
    ("Agua",None,None),("Wifi (Telecentro)",38940,None),("Limpieza",198000,63000),
    ("Paseadora Antonio",98000,None),("Netflix / Streaming",22648,None),
    ("Pilates",None,None),("ABL / ARBA",None,None),("Seguro hogar",None,None),
]
rr=15
for nom,n,p in ref:
    ing.cell(rr,1,nom);
    ing.cell(rr,2, n if n else None); ing.cell(rr,3, p if p else None)
    style(ing.cell(rr,1), border=border_all)
    style(ing.cell(rr,2), border=border_all, bg=AMAR, numfmt=MONEY, align="right")
    style(ing.cell(rr,3), border=border_all, bg=AMAR, numfmt=MONEY, align="right")
    rr+=1
ing.merge_cells(f"A{rr+1}:E{rr+2}")
ing.cell(rr+1,1,"Todas las celdas amarillas son editables. Si pisás una fórmula perdés el cálculo automático de esa celda, pero el archivo no se rompe.")
style(ing.cell(rr+1,1), italic=True, wrap=True, color="595959")

# ======================================================================
# FUNCION: crear hoja de mes
# ======================================================================
MESES = ["JUNIO","JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]

# data real (solo Jun y Jul). Estructura:
# fijos: dict concepto -> (monto, medio)   medio: "AMEX Loli"/"Transferencia"/""
# amex : lista de (detalle, categoria, casa, monto, medio)
DATA = {
 "JUNIO": {
   "fijos": {
     "Expensas Núñez": (None,"AMEX Loli"),
     "Expensas Pilar": (398000,"AMEX Loli"),
     "Luz Núñez": (52392,"AMEX Loli"),
     "Luz Pilar": (None,"AMEX Loli"),
     "Wifi Núñez": (38940,"AMEX Loli"),
     "Wifi Pilar": (None,"AMEX Loli"),
     "Paseadora Antonio": (None,"Transferencia"),
     "Limpieza Núñez": (None,"Transferencia"),
     "Limpieza Pilar": (None,"Transferencia"),
     "Netflix": (22648,"AMEX Loli"),
     "Pilates": (None,"Transferencia"),
   },
   "amex": [
     ("Hamburguesería Colegiales","Restaurant","Ambos",47000,"AMEX Loli"),
   ],
   "aporte_loli":1000000, "aporte_tito":1500000, "sobro_prev":0,
 },
 "JULIO": {
   "fijos": {
     "Expensas Núñez": (415000,"AMEX Loli"),
     "Expensas Pilar": (398000,"AMEX Loli"),
     "Luz Núñez": (21318,"AMEX Loli"),
     "Luz Pilar": (24579,"AMEX Loli"),
     "Wifi Núñez": (44898,"AMEX Loli"),
     "Wifi Pilar": (45238,"AMEX Loli"),
     "Paseadora Antonio": (98000,"Transferencia"),
     "Limpieza Núñez": (198000,"Transferencia"),
     "Limpieza Pilar": (63000,"Transferencia"),
     "Netflix": (22648,"AMEX Loli"),
     "Pilates": (None,"Transferencia"),
   },
   "amex": [
     ("Merpago Delfin","Supermercado","Núñez",23500,"AMEX Loli"),
     ("Groovypets","Veterinaria","Ambos",50900,"AMEX Loli"),
     ("Merpago-h2soap","Otros","Ambos",14944,"AMEX Loli"),
     ("Groovypets","Veterinaria","Ambos",44500,"AMEX Loli"),
     ("Rappi","Delivery","Núñez",50897,"AMEX Loli"),
     ("Super Manzanares","Supermercado","Núñez",25000,"AMEX Loli"),
     ("Pedidos Ya market","Supermercado","Núñez",34800,"AMEX Loli"),
     ("Rappi Super","Supermercado","Núñez",156500,"AMEX Loli"),
     ("La Noire Café + Razo-Razo","Cafecito","Ambos",48900,"AMEX Loli"),
     ("Merpago Delfin + Super Paroissien","Supermercado","Núñez",92000,"AMEX Loli"),
     ("Solileb","Cafecito","Ambos",74000,"AMEX Loli"),
     ("Rappi","Delivery","Ambos",22162,"AMEX Loli"),
     ("JETSMART","Viaje","Ambos",743000,"AMEX Loli"),
     ("Rappi","Delivery","Ambos",17192,"AMEX Loli"),
     ("Rappi","Delivery","Ambos",37492,"AMEX Loli"),
     ("Parmegiano","Restaurant","Ambos",46458,"AMEX Loli"),
     ("Santicheese","Restaurant","Ambos",109800,"AMEX Loli"),
     ("dlorappi","Delivery","Ambos",6490,"AMEX Loli"),
     ("Merpago-Delfin","Supermercado","Ambos",30246,"AMEX Loli"),
     ("Vacalin","Supermercado","Ambos",33100,"AMEX Loli"),
     ("Rappi","Delivery","Ambos",32554,"AMEX Loli"),
     ("Rappi","Delivery","Ambos",64450,"AMEX Loli"),
     ("Demente","Supermercado","Ambos",14000,"AMEX Loli"),
     ("Evelia","Restaurant","Ambos",160400,"AMEX Loli"),
   ],
   "aporte_loli":1000000, "aporte_tito":2000000, "sobro_prev":None,  # linkea a Junio
 },
}

FIJOS_CONCEPTOS = ["Expensas Núñez","Expensas Pilar","Luz Núñez","Luz Pilar",
    "Wifi Núñez","Wifi Pilar","Paseadora Antonio","Limpieza Núñez","Limpieza Pilar",
    "Netflix","Pilates"]

month_sheets = {}

def build_month(name, prev_name):
    ws = wb.create_sheet(name)
    for col,w in zip("ABCDEFG",[30,14,12,14,16,14,14]): ws.column_dimensions[col].width = w
    d = DATA.get(name, {"fijos":{},"amex":[],"aporte_loli":0,"aporte_tito":0,"sobro_prev":None})

    # ---- Titulo
    ws.merge_cells("A1:G1")
    ws["A1"]=f"{name} 2026"
    style(ws["A1"], bold=True, size=18, color=BLANCO, bg=AZUL, align="center")
    ws.row_dimensions[1].height=30

    # =========== PANEL DE CONTROL ===========
    ws.merge_cells("A3:G3")
    ws["A3"]="CONTROL DEL MES"
    style(ws["A3"], bold=True, size=12, color=BLANCO, bg=AZUL2, align="center")

    # etiquetas fila 4, valores fila 5 (editables amarillo)
    labels = ["Aporte Loli","Aporte Tito","Sobró del mes anterior","= Total disponible"]
    for i,l in enumerate(labels):
        style(ws.cell(4,1+i,l), bold=True, color=BLANCO, bg=GRIS_HDR, align="center", border=border_all, wrap=True)
    ws.cell(5,1, d["aporte_loli"])
    ws.cell(5,2, d["aporte_tito"])
    # sobro_prev: link al mes anterior si existe
    if prev_name is None:
        ws.cell(5,3, d.get("sobro_prev") or 0)
    else:
        ws.cell(5,3, f"='{prev_name}'!$G$5")   # saldo a favor del mes anterior
    ws.cell(5,4, "=A5+B5+C5")
    style(ws.cell(5,1), bg=AMAR, border=border_box, numfmt=MONEYAR, align="right", bold=True)
    style(ws.cell(5,2), bg=AMAR, border=border_box, numfmt=MONEYAR, align="right", bold=True)
    style(ws.cell(5,3), bg=AMAR, border=border_box, numfmt=MONEYAR, align="right", bold=True)
    style(ws.cell(5,4), bg=CELESTE2, border=border_box, numfmt=MONEYAR, align="right", bold=True)
    ws.row_dimensions[4].height=26

    # resultado: total gastado / sobro / falto  (fila 4-5 columnas E,F,G)
    for i,l in zip(range(4,7),["Total gastado","SOBRÓ (saldo a favor →)","FALTÓ"]):
        pass
    style(ws.cell(4,5,"Total gastado"), bold=True, color=BLANCO, bg=GRIS_HDR, align="center", border=border_all, wrap=True)
    style(ws.cell(4,6,"FALTÓ"), bold=True, color=BLANCO, bg=ROJO_TXT, align="center", border=border_all, wrap=True)
    style(ws.cell(4,7,"SOBRÓ (pasa →)"), bold=True, color=BLANCO, bg=VERDE_TXT, align="center", border=border_all, wrap=True)
    # Total gastado = subtotales (se completan mas abajo con refs). Placeholder set later.
    # SOBRO = MAX(0, disponible - gastado) ; FALTO = MAX(0, gastado - disponible)
    ws.cell(5,5, None)  # se setea luego = total gastado
    ws.cell(6,6, None)
    # colocamos formulas cuando conozcamos filas de subtotales
    style(ws.cell(5,5), bg=CELESTE2, border=border_box, numfmt=MONEYAR, align="right", bold=True)
    ws.cell(5,6, "=MAX(0,E5-D5)")   # FALTO
    ws.cell(5,7, "=MAX(0,D5-E5)")   # SOBRO
    style(ws.cell(5,6), bg=ROJO, border=border_box, numfmt=MONEYAR, align="right", bold=True, color=ROJO_TXT)
    style(ws.cell(5,7), bg=VERDE, border=border_box, numfmt=MONEYAR, align="right", bold=True, color=VERDE_TXT)

    # =========== TABLA GASTOS FIJOS ===========
    rowf = 8
    ws.merge_cells(f"A{rowf}:G{rowf}")
    ws.cell(rowf,1,"GASTOS FIJOS   ·   monto y medio de pago")
    style(ws.cell(rowf,1), bold=True, size=12, color=BLANCO, bg=GRIS_HDR, align="left")
    rowf+=1
    for i,h in enumerate(["Concepto","Monto","Medio de pago"]):
        style(ws.cell(rowf,1+i,h), bold=True, color=BLANCO, bg=AZUL2, align="center", border=border_all)
    ws.merge_cells(f"C{rowf}:D{rowf}")   # concepto A, monto B, medio C:D
    # reorganizo: A concepto, B monto, C medio
    ws.cell(rowf,1,"Concepto"); ws.cell(rowf,2,"Monto"); ws.cell(rowf,3,"Medio de pago")
    ws.unmerge_cells(f"C{rowf}:D{rowf}")
    for i,h in enumerate(["Concepto","Monto","Medio de pago"]):
        style(ws.cell(rowf,1+i,h), bold=True, color=BLANCO, bg=AZUL2, align="center", border=border_all)
    fijo_start = rowf+1
    r = fijo_start
    for concepto in FIJOS_CONCEPTOS:
        monto, medio = d["fijos"].get(concepto,(None,""))
        ws.cell(r,1,concepto)
        ws.cell(r,2, monto)
        ws.cell(r,3, medio)
        style(ws.cell(r,1), border=border_all)
        style(ws.cell(r,2), border=border_all, bg=AMAR, numfmt=MONEY, align="right")
        style(ws.cell(r,3), border=border_all, align="center")
        r+=1
    # filas en blanco "por si me olvido"
    for _ in range(4):
        ws.cell(r,1,None)
        style(ws.cell(r,1), border=border_all, bg=BLANCO)
        style(ws.cell(r,2), border=border_all, bg=AMAR, numfmt=MONEY, align="right")
        style(ws.cell(r,3), border=border_all, align="center")
        r+=1
    fijo_end = r-1
    # subtotal fijos
    ws.cell(r,1,"Subtotal fijos")
    ws.cell(r,2, f"=SUM(B{fijo_start}:B{fijo_end})")
    style(ws.cell(r,1), bold=True, bg=CELESTE, border=border_all)
    style(ws.cell(r,2), bold=True, bg=CELESTE, border=border_all, numfmt=MONEYAR, align="right")
    style(ws.cell(r,3), bg=CELESTE, border=border_all)
    sub_fijos_row = r
    r+=2

    # =========== TABLA GASTOS AMEX (variables/categorias) ===========
    ws.merge_cells(f"A{r}:G{r}")
    ws.cell(r,1,"GASTOS AMEX / TARJETA   ·   por categoría (se acomoda solo)")
    style(ws.cell(r,1), bold=True, size=12, color=BLANCO, bg=GRIS_HDR, align="left")
    r+=1
    amex_hdr_row = r
    for i,h in enumerate(["Detalle / Comercio","Categoría (auto)","Casa","Monto","Medio de pago"]):
        style(ws.cell(r,1+i,h), bold=True, color=BLANCO, bg=AZUL2, align="center", border=border_all)
    r+=1
    amex_start = r
    filas_amex = d["amex"]
    N_AMEX = max(len(filas_amex)+12, 30)   # espacio de sobra para pegar el resumen
    for idx in range(N_AMEX):
        if idx < len(filas_amex):
            det,cat,casa,monto,medio = filas_amex[idx]
            ws.cell(r,1,det)
            # categoria: formula auto con override manual -> ponemos formula
            ws.cell(r,2, f'=IF($A{r}="","",IFERROR(LOOKUP(2,1/((ISNUMBER(SEARCH({MAP_RANGE_KW},$A{r})))*({MAP_RANGE_KW}<>"")),{MAP_RANGE_CAT}),"Otros"))')
            ws.cell(r,3,casa)
            ws.cell(r,4,monto)
            ws.cell(r,5,medio)
        else:
            ws.cell(r,1,None)
            ws.cell(r,2, f'=IF($A{r}="","",IFERROR(LOOKUP(2,1/((ISNUMBER(SEARCH({MAP_RANGE_KW},$A{r})))*({MAP_RANGE_KW}<>"")),{MAP_RANGE_CAT}),"Otros"))')
        style(ws.cell(r,1), border=border_all)
        style(ws.cell(r,2), border=border_all, align="center", color="1F4E78")
        style(ws.cell(r,3), border=border_all, align="center")
        style(ws.cell(r,4), border=border_all, bg=AMAR, numfmt=MONEY, align="right")
        style(ws.cell(r,5), border=border_all, align="center")
        r+=1
    amex_end = r-1
    ws.cell(r,1,"Subtotal AMEX / tarjeta")
    ws.cell(r,4, f"=SUM(D{amex_start}:D{amex_end})")
    style(ws.cell(r,1), bold=True, bg=CELESTE, border=border_all)
    style(ws.cell(r,2), bg=CELESTE, border=border_all)
    style(ws.cell(r,3), bg=CELESTE, border=border_all)
    style(ws.cell(r,4), bold=True, bg=CELESTE, border=border_all, numfmt=MONEYAR, align="right")
    style(ws.cell(r,5), bg=CELESTE, border=border_all)
    sub_amex_row = r
    r+=1
    # nota pegar resumen
    ws.merge_cells(f"A{r}:G{r}")
    ws.cell(r,1,"➤ Pegá acá el resumen de la AMEX (Detalle + Monto). La categoría se completa sola según el mapa de la hoja CONFIG. No hace falta crear otro archivo.")
    style(ws.cell(r,1), italic=True, color="595959", wrap=True)
    ws.row_dimensions[r].height=28
    r+=2

    # ---- ahora completo Total gastado del panel
    ws.cell(5,5, f"=B{sub_fijos_row}+D{sub_amex_row}")

    # =========== RESUMEN ===========
    ws.merge_cells(f"A{r}:G{r}")
    ws.cell(r,1,"RESUMEN · EN QUÉ SE GASTÓ MÁS")
    style(ws.cell(r,1), bold=True, size=12, color=BLANCO, bg=AZUL, align="center")
    r+=1
    # Por seccion
    ws.cell(r,1,"Por sección"); style(ws.cell(r,1), bold=True, bg=CELESTE, border=border_all)
    for i,h in enumerate(["Monto","% del total"],start=1):
        style(ws.cell(r,1+i,h), bold=True, bg=CELESTE, border=border_all, align="center")
    style(ws.cell(r,2), bold=True, bg=CELESTE, border=border_all)  # spacer col? keep simple
    r+=1
    total_mes_ref = f"$B${sub_fijos_row}+$D${sub_amex_row}"
    sec_rows=[]
    for nom, ref in [("Gastos fijos", f"$B${sub_fijos_row}"),
                     ("Gastos AMEX / tarjeta", f"$D${sub_amex_row}")]:
        ws.cell(r,1,nom); style(ws.cell(r,1), border=border_all)
        ws.cell(r,2, f"={ref}"); style(ws.cell(r,2), border=border_all, numfmt=MONEYAR, align="right")
        ws.cell(r,3, f"=IF(({total_mes_ref})=0,0,({ref})/({total_mes_ref}))")
        style(ws.cell(r,3), border=border_all, numfmt=PCT, align="center")
        sec_rows.append(r); r+=1
    ws.cell(r,1,"TOTAL"); style(ws.cell(r,1), bold=True, bg=CELESTE2, border=border_all)
    ws.cell(r,2, f"={total_mes_ref}"); style(ws.cell(r,2), bold=True, bg=CELESTE2, border=border_all, numfmt=MONEYAR, align="right")
    ws.cell(r,3, 1 if False else "=IF(B{0}=0,0,1)".format(r));
    ws.cell(r,3,"=IF(("+total_mes_ref+")=0,0,1)")
    style(ws.cell(r,3), bold=True, bg=CELESTE2, border=border_all, numfmt=PCT, align="center")
    r+=2

    # Por categoria (AMEX)
    ws.cell(r,1,"Por categoría (AMEX / tarjeta)");
    style(ws.cell(r,1), bold=True, bg=CELESTE, border=border_all)
    style(ws.cell(r,2,"Monto"), bold=True, bg=CELESTE, border=border_all, align="center")
    style(ws.cell(r,3,"% AMEX"), bold=True, bg=CELESTE, border=border_all, align="center")
    r+=1
    cat_first=r
    amex_cat_col = f"$B${amex_start}:$B${amex_end}"
    amex_mon_col = f"$D${amex_start}:$D${amex_end}"
    for ct in CATS:
        ws.cell(r,1,ct); style(ws.cell(r,1), border=border_all)
        ws.cell(r,2, f'=SUMIF({amex_cat_col},A{r},{amex_mon_col})')
        style(ws.cell(r,2), border=border_all, numfmt=MONEY, align="right")
        ws.cell(r,3, f'=IF($D${sub_amex_row}=0,0,B{r}/$D${sub_amex_row})')
        style(ws.cell(r,3), border=border_all, numfmt=PCT, align="center")
        r+=1
    cat_last=r-1
    # Top gasto
    ws.cell(r,1,"Top categoría"); style(ws.cell(r,1), bold=True, bg=VERDE, border=border_all, color=VERDE_TXT)
    ws.cell(r,2, f'=IFERROR(INDEX(A{cat_first}:A{cat_last},MATCH(MAX(B{cat_first}:B{cat_last}),B{cat_first}:B{cat_last},0)),"-")')
    style(ws.cell(r,2), bold=True, bg=VERDE, border=border_all, align="center", color=VERDE_TXT)
    ws.cell(r,3, f'=MAX(B{cat_first}:B{cat_last})')
    style(ws.cell(r,3), bold=True, bg=VERDE, border=border_all, numfmt=MONEY, align="right", color=VERDE_TXT)
    r+=2

    # =========== SUGERIDO MES SIGUIENTE (gris) ===========
    ws.merge_cells(f"A{r}:G{r}")
    ws.cell(r,1,"SUGERIDO PARA EL MES SIGUIENTE  ·  según lo gastado este mes")
    style(ws.cell(r,1), bold=True, size=12, color="000000", bg=GRIS_MED, align="center")
    r+=1
    sug_hdr=r
    for i,h in enumerate(["Base (gasto del mes)","% Loli","% Tito","Sugerido Loli","Sugerido Tito"]):
        style(ws.cell(r,1+i,h), bold=True, bg=GRIS_CL, align="center", border=border_all, wrap=True)
    r+=1
    ws.cell(r,1, f"={total_mes_ref}")
    ws.cell(r,2, "=PCT_LOLI")
    ws.cell(r,3, "=PCT_TITO")
    ws.cell(r,4, f"=A{r}*B{r}")
    ws.cell(r,5, f"=A{r}*C{r}")
    style(ws.cell(r,1), bg=GRIS_CL, border=border_all, numfmt=MONEYAR, align="right")
    style(ws.cell(r,2), bg=GRIS_CL, border=border_all, numfmt=PCT, align="center")
    style(ws.cell(r,3), bg=GRIS_CL, border=border_all, numfmt=PCT, align="center")
    style(ws.cell(r,4), bg=GRIS_CL, border=border_all, numfmt=MONEYAR, align="right", bold=True)
    style(ws.cell(r,5), bg=GRIS_CL, border=border_all, numfmt=MONEYAR, align="right", bold=True)
    r+=1
    ws.merge_cells(f"A{r}:G{r}")
    ws.cell(r,1,"Idea: para no quedar cortos, cada uno aporta aprox. lo que le tocó este mes según el split 40/60. Si querés margen, subí un 10-15%.")
    style(ws.cell(r,1), italic=True, color="595959", wrap=True)
    ws.row_dimensions[r].height=26

    # ---- Data validations
    dv_medio = DataValidation(type="list", formula1='"AMEX Loli,AMEX Tito,Transferencia,Efectivo"', allow_blank=True)
    dv_casa  = DataValidation(type="list", formula1='"Núñez,Pilar,Ambos"', allow_blank=True)
    ws.add_data_validation(dv_medio); ws.add_data_validation(dv_casa)
    dv_medio.add(f"C{fijo_start}:C{fijo_end}")
    dv_medio.add(f"E{amex_start}:E{amex_end}")
    dv_casa.add(f"C{amex_start}:C{amex_end}")

    ws.sheet_view.showGridLines = False
    month_sheets[name]=ws
    return ws

prev=None
for m in MESES:
    build_month(m, prev)
    prev=m

# ======================================================================
# RESUMEN ANUAL
# ======================================================================
ra = wb.create_sheet("RESUMEN ANUAL")
for col,w in zip("ABCDEFGHIJ",[16,12,12,12,12,12,12,12,12,14]): ra.column_dimensions[col].width=w
ra.merge_cells("A1:J1"); ra["A1"]="RESUMEN ANUAL · 2026"
style(ra["A1"], bold=True, size=16, color=BLANCO, bg=AZUL, align="center")
ra.row_dimensions[1].height=28
cols_meses = MESES  # Junio..Diciembre
# encabezado
ra.cell(3,1,"Concepto"); style(ra.cell(3,1), bold=True, color=BLANCO, bg=GRIS_HDR, align="center", border=border_all)
for i,m in enumerate(cols_meses):
    style(ra.cell(3,2+i,m.title()[:3]), bold=True, color=BLANCO, bg=GRIS_HDR, align="center", border=border_all)
style(ra.cell(3,2+len(cols_meses),"TOTAL"), bold=True, color=BLANCO, bg=AZUL2, align="center", border=border_all)
tot_col = 2+len(cols_meses)

def fila_ra(row, label, cellref, fmt=MONEYAR, bg=None):
    ra.cell(row,1,label); style(ra.cell(row,1), bold=True, border=border_all, bg=bg or CELESTE)
    for i,m in enumerate(cols_meses):
        ra.cell(row,2+i, f"='{m}'!{cellref}")
        style(ra.cell(row,2+i), border=border_all, numfmt=fmt, align="right", bg=bg)
    c0=get_column_letter(2); c1=get_column_letter(1+len(cols_meses))
    ra.cell(row,tot_col, f"=SUM({c0}{row}:{c1}{row})")
    style(ra.cell(row,tot_col), bold=True, border=border_all, numfmt=fmt, align="right", bg=CELESTE2)

fila_ra(4,"Total gastado","$E$5")
fila_ra(5,"Aporte Loli","$A$5")
fila_ra(6,"Aporte Tito","$B$5")
fila_ra(7,"Sobró (saldo a favor)","$G$5", bg=VERDE)
fila_ra(8,"Faltó","$F$5", bg=ROJO)
ra.sheet_view.showGridLines=False

# ---- orden de hojas
order = ["INGRESOS"]+MESES+["RESUMEN ANUAL","CONFIG"]
wb._sheets.sort(key=lambda s: order.index(s.title) if s.title in order else 99)
ing.sheet_view.showGridLines=False
cfg.sheet_view.showGridLines=False

out = "/home/user/FINANZAS_LOLI_TITO/FINANZAS_LOLI_TITO.xlsx"
wb.save(out)
print("Guardado:", out)
print("Hojas:", [s.title for s in wb._sheets])
