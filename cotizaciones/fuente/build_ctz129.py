from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import os
OUT=os.path.join(os.path.dirname(os.path.abspath(__file__)),"..","CTZ-129.2026 Pabellon China Expocruz 2027 - DESPIECE.xlsx")
EST="ESTIMADO"
S=dict(cen="COTIZ-104.2026REV03 – bastidor cenefa 191,90/m²",corp="COTIZ-104.2026REV03 – bastidor c/cuerpo 275,00/m²",
 adh="CTZ-123/115.2026 – adhesivo 170,00/m²",adhm="COTIZ-092.2026 – adhesivo credenza 160,15/pza",
 alf="COTIZ-081.2026 – alfombra 168,67/m²",muro="COTIZ-103.2026REV01 – muro con cuerpo 546,43/m²",
 tar="COTIZ-069.2026 – tarima 1.003,09/m²",let="COTIZ-092.2026 – letras 3D 4.735,71",
 foam="COTIZ-103.2026 – letras foam 2.822,30",mmdf="COTIZ-103.2026REV01 – mobiliario MDF 6.500,00",
 tot="COTIZ-092.2026 – tótem bastidor 738,82",tec="COTIZ-081.2026 – 3.500 / 5 días, x 10 días",
 bas="Derivado CTZ-123.2026 (218) − CTZ-067.2026 lona (122,88) + imp. = 113,10/m²",
 dw="CTZ-126.2026 – muro drywall doble cara 272,67 + imp. (x1,1905) = 324,61/m²",
 dry="Derivado CTZ-126.2026 (muro drywall doble cara 272,67 / 2 caras) = 136,34/m² por cara",
 mdf="Derivado CTZ-123.2026 (326 − 218) + imp. = 128,57/m²")
PB,PM,PA=113.10,128.57,170.00
PD=136.34
PW=324.61
INC={S["cen"]:"INCLUYE ESTRUCTURA METÁLICA, LONA CON IMPRESIÓN DIGITAL FULL COLOR E INSTALACIÓN.",
 S["corp"]:"INCLUYE ESTRUCTURA METÁLICA CON CUERPO, LONA CON IMPRESIÓN DIGITAL FULL COLOR E INSTALACIÓN.",
 S["muro"]:"INCLUYE ESTRUCTURA METÁLICA, TABLERO MDF, PINTURA E INSTALACIÓN.",
 S["adh"]:"INCLUYE IMPRESIÓN DIGITAL FULL COLOR Y COLOCADO.",S["adhm"]:"INCLUYE IMPRESIÓN Y COLOCADO SOBRE EL MUEBLE QUE PROVEE LA FERIA.",
 S["alf"]:"INCLUYE PEGAMENTO DE CONTACTO Y COLOCADO.",S["tar"]:"INCLUYE MATERIAL, FABRICACIÓN E INSTALACIÓN.",
 S["let"]:"INCLUYE MATERIAL, CORTE LÁSER E INSTALACIÓN.",S["foam"]:"INCLUYE MATERIAL, CORTE LÁSER E INSTALACIÓN."}
def P(d,L,A,F,N,u,pu,src,obs=""): return ("P",d,L,A,F,N,u,pu,src,obs)
def B(d,L,A,N,m=1,a=1): return ("B",d,L,A,N,m,a)
def W(d,L,A,N,m=1,a=1): return ("W",d,L,A,N,m,a)
# (título, categoría comparativo, ítems)
G=[
("STAND ESTÁNDAR 9 m² (26 STANDS: 22 UNA CARA ABIERTA + 4 DOS CARAS)","STANDS B",[
 P("Top banner (cenefa) 2,92x0,85, bastidor con lona impresa una cara",2.92,0.85,1,30,"M2",191.90,S["cen"],"22 x 1 + 4 x 2 según brief CCPIT"),
 P("Fascia name (franja nombre) 2,92x0,195, bastidor con lona impresa una cara",2.92,0.195,1,30,"M2",191.90,S["cen"],"22 x 1 + 4 x 2"),
 P("Gráfica para info counter 0,895x0,625",None,None,None,26,"Pza",160.15,S["adhm"],"Counter lo provee la feria")]),
("STAND ESTÁNDAR 18 m² (5 STANDS: 1 UNA CARA + 3 DOS CARAS + 1 TRES CARAS)","STANDS B",[
 P("Top banner (cenefa) 2,92x0,85, bastidor con lona impresa una cara",2.92,0.85,1,15,"M2",191.90,S["cen"],"1 x 2 + 3 x 3 + 1 x 4 (Julong)"),
 P("Fascia name (franja nombre) 2,92x0,195, bastidor con lona impresa una cara",2.92,0.195,1,15,"M2",191.90,S["cen"],""),
 P("Gráfica para info counter 0,895x0,625",None,None,None,5,"Pza",160.15,S["adhm"],"")]),
("STAND ESTÁNDAR 36 m² BYD (1 STAND)","STANDS B",[
 P("Top banner (cenefa) 2,92x0,85, bastidor con lona impresa una cara",2.92,0.85,1,2,"M2",191.90,S["cen"],"O reemplazar por logo BYD"),
 P("Gráfica para info counter 0,895x0,625",None,None,None,1,"Pza",160.15,S["adhm"],"")]),
("STAND A-01 HIKVISION 4,00 x 7,50 m","STANDS A",[
 W("Bastidor 1",7.50,4.00,1),W("Bastidor 2",4.00,4.00,1),W("Bastidor 3",2.09,4.00,1),W("Bastidor 4 (cenefa)",7.39,1.30,2),
 W("Bastidor 5 (cenefa)",1.69,1.30,3),W("Bastidor 6",2.44,4.00,1),W("Bastidor 7",1.69,4.00,1),W("Bastidor 8",1.69,0.40,3),W("Bastidor 9",7.50,0.40,1),
 ("ADH",""),
 P("Puerta de depósito MDF con chapa 0,80x2,10",None,None,None,1,"Pza",900,EST),
 P("Letras 3D \"HIKVISION\" 2,22x0,32",None,None,None,1,"Pza",4735.71,S["let"]),
 P("Counter de información en L, laminado, con puertas y chapa 2,00x1,35x0,90 (diseño del expositor)",None,None,None,1,"Pza",6500,S["mmdf"],"Mobiliario especial"),
 P("Vitrina redonda escalonada Ø0,80 h=1,25, pintura al duco (diseño del expositor)",None,None,None,1,"Pza",2500,EST,"Mobiliario especial")]),
("STAND A-02 GINLONG – SOLIS 4,00 x 7,50 m","STANDS A",[
 W("Bastidor 1",7.50,4.00,1),W("Bastidor 2",3.95,4.00,1),W("Bastidor 3 (tabique, adhesivo 2 caras)",1.50,4.00,1,1,2),W("Bastidor 4 (tabique, adhesivo 2 caras)",2.14,4.00,1,1,2),
 ("ADH"," (MUROS 3 Y 4 A DOS CARAS)"),
 P("Adhesivo para counter 2,07x0,78",None,None,None,1,"Pza",160.15,S["adhm"],"Counter lo provee la feria"),
 P("Display de inversores: estructura de acero, MDF con adhesivo 2,80x1,20x1,91 (diseño del expositor)",None,None,None,1,"Pza",9500,EST,"Mobiliario especial")]),
("STAND A-03 WEIQIAO (212) 4,00 x 7,50 m","STANDS A",[
 W("Bastidor 1",7.50,4.00,1),W("Bastidor 2",4.00,4.00,1),("ADH",""),
 P("Adhesivo números \"212\" para counter 1,20x0,91",None,None,None,1,"Pza",160.15,S["adhm"],"Vehículo y counter no incluidos")]),
("STAND A-04 KONKA 4,00 x 7,50 m","STANDS A",[
 W("Bastidor 1",7.50,4.00,1),W("Bastidor 2",4.00,4.00,1),W("Bastidor 3 corpóreo (A 4,00 · B 2,13 · C 4,36; área según volúmenes)",10.49,1.00,1),("ADH",""),
 P("Adhesivo para credenzas y counter",None,None,None,4,"Pza",160.15,S["adhm"],"Mobiliario lo provee la feria")]),
("STAND PAÍS · PÓRTICO \"Y\" (ARCOS CENTRALES)","STAND PAÍS",[
 P("Tramo diagonal posterior, corpóreo sección 0,77x0,30, lona 3 caras",4.50,1.84,1,2,"M2",275.00,S["corp"],"A = desarrollo de sección"),
 P("Tramo recto central, corpóreo, lona 3 caras",5.00,1.84,1,2,"M2",275.00,S["corp"]),
 P("Tramo diagonal frontal, corpóreo, lona 3 caras",4.96,1.84,1,2,"M2",275.00,S["corp"]),
 P("Viga transversal posterior, bastidor con lona 3 caras",6.42,1.00,1,1,"M2",275.00,S["corp"]),
 P("Viga transversal frontal, bastidor con lona 3 caras",5.67,1.00,1,1,"M2",275.00,S["corp"]),
 P("Placa corpórea trapezoidal soporte de letras, lona 2 caras",2.14,0.40,2,1,"M2",275.00,S["corp"],"Medida a escala"),
 P("Letras 3D \"CHINA\" cara frontal 1,40x0,30",None,None,None,1,"Pza",2822.30,S["foam"],"Medida a escala"),
 P("Letras 3D \"CHINA\" cara posterior 1,15x0,30",None,None,None,1,"Pza",2822.30,S["foam"],"Medida a escala")]),
("STAND PAÍS · ESCENARIO (ESTRUCTURA DE PANTALLA) 6,50 x 3,20 x 1,00 m","STAND PAÍS",[
 B("Cara posterior (la cara frontal la cubre la pantalla LED)",6.50,3.20,1,1,0),B("Laterales",1.00,3.20,2,1,0),("MDF",)]),
("STAND PAÍS · TÓTEMS CENTRALES 1,90 x 3,20 m (2 UND)","STAND PAÍS",[
 B("Caras",1.90,3.20,2,1,0),B("Cantos",0.30,3.20,2,1,0),("MDF",)]),
("STAND PAÍS · TÓTEMS DE FONDO 3,00 x 2,50 m (4 UND)","STAND PAÍS",[
 B("Caras",3.00,2.50,4,1,0),B("Cantos",0.20,2.50,4,1,0),("MDF",)]),
("STAND PAÍS · TARIMA Y PISO","STAND PAÍS",[
 P("Tarima h=20 cm, tubin 20x20, fenólico 18 mm, forrada con alfombra",6.50,2.00,1,1,"M2",1003.09,S["tar"])]),
("PÓRTICO DE INGRESO PRINCIPAL 6,65 x 4,26 m","PÓRTICO DE INGRESO",[
 P("Panel inclinado izquierdo, estructura metálica con cuerpo forrada con lona impresa (rojo + patrón olas)",1.25,3.50,2,1,"M2",275.00,S["corp"],"2 caras"),
 P("Panel vertical esbelto, estructura metálica con cuerpo forrada con lona impresa",0.38,4.26,2,1,"M2",275.00,S["corp"],"2 caras"),
 P("Pata izquierda del marco, estructura metálica con cuerpo forrada con lona impresa",0.80,4.12,2,1,"M2",275.00,S["corp"],"2 caras"),
 P("Dintel, estructura metálica con cuerpo forrada con lona impresa",5.17,1.00,2,1,"M2",275.00,S["corp"],"2 caras"),
 P("Pata derecha, estructura metálica con cuerpo forrada con lona impresa",1.00,4.00,2,1,"M2",275.00,S["corp"],"2 caras"),
 P("Placa frontal \"CHINA\", estructura metálica con cuerpo forrada con lona impresa",3.40,1.00,2,1,"M2",275.00,S["corp"],"2 caras"),
 P("Placa inclinada pequeña, estructura metálica con cuerpo forrada con lona impresa",0.50,1.00,2,1,"M2",275.00,S["corp"],"2 caras"),
 P("Cantos del marco principal, e=0,20",28.18,0.20,1,1,"M2",275.00,S["corp"],"Perímetro marco en U"),
 P("Cantos panel vertical esbelto, e=0,20",9.28,0.20,1,1,"M2",275.00,S["corp"]),
 P("Cantos panel inclinado, e=0,10",9.50,0.10,1,1,"M2",275.00,S["corp"]),
 P("Cantos placa \"CHINA\", e=0,10",8.80,0.10,1,1,"M2",275.00,S["corp"]),
 P("Cantos placa inclinada pequeña, e=0,10",3.00,0.10,1,1,"M2",275.00,S["corp"]),
 P("Letras 3D \"CHINA\" blancas 1,90x0,35",None,None,None,1,"Pza",4735.71,S["let"]),
 P("Números \"2027\" en foam adhesivado",None,None,None,1,"Pza",2822.30,S["foam"]),
 P("Chevrón decorativo 1,00x1,23, estructura metálica forrada con foam, acabado patrón olas",None,None,None,2,"Pza",850,EST)]),
("MURAL DECORATIVO (PARED DE INGRESO)","MURAL DECORATIVO",[
 P("Bastidor corpóreo panel cuadrado, lona patrón olas, marco 0,20",3.50,3.50,1,1,"M2",275.00,S["corp"]),
 P("Bastidor corpóreo panel horizontal, lona patrón olas, marco 0,20",7.00,1.70,1,1,"M2",275.00,S["corp"]),
 P("Barra corpórea superior (soporte letras) 0,20x0,20",5.20,0.60,1,1,"M2",275.00,S["corp"],"A = desarrollo 3 caras"),
 P("Marco corpóreo paisaje 0,20x0,20",11.10,0.60,1,1,"M2",275.00,S["corp"]),
 P("Silueta montañas MDF corte CNC con adhesivo",3.50,0.50,1,1,"M2",546.43,S["muro"]),
 P("Círculo corpóreo Ø1,10 con aro",None,None,None,1,"Pza",900,EST),
 P("Letras 3D \"CHINA\" 2,12x0,45",None,None,None,2,"Pza",4735.71,S["let"],"2 juegos")]),
("BANDEROLAS (6 UND)","SEÑALÉTICA",[
 P("Mástil metálico h=2,50 m con base circular",None,None,None,6,"Pza",450,EST),
 P("Bandera lona impresa doble cara 0,60x1,52",0.60,1.52,1,6,"M2",191.90,S["cen"])]),
("TÓTEMS DIRECCIONALES 0,80 x 1,80 m (4 UND)","SEÑALÉTICA",[
 P("Tótem bastidor con lona 2 caras y base",None,None,None,4,"Pza",738.82,S["tot"])]),
("LOGÍSTICA, DESMONTAJE Y ACOMPAÑAMIENTO","GENERAL",[
 P("Logística y transporte de materiales",None,None,None,1,"Glb",6000,EST),
 P("Desmontaje y retiro de materiales",None,None,None,1,"Glb",8000,EST,"Ref. COTIZ-107.2026"),
 P("Acompañamiento técnico durante feria (10 días)",None,None,None,1,"Glb",7000,S["tec"])]),
]

_G={t:(cat,its) for t,cat,its in G}
def sub(t,name=None): return ("SUB",name or t,_G[t][1])
T=[
("STAND PAÍS 10,50 x 15,00 m","STAND PAÍS",[sub('STAND PAÍS · PÓRTICO "Y" (ARCOS CENTRALES)','PÓRTICO "Y" (ARCOS CENTRALES)'),
   sub("STAND PAÍS · ESCENARIO (ESTRUCTURA DE PANTALLA) 6,50 x 3,20 x 1,00 m","ESCENARIO (ESTRUCTURA DE PANTALLA) 6,50 x 3,20 x 1,00 m"),
   sub("STAND PAÍS · TÓTEMS CENTRALES 1,90 x 3,20 m (2 UND)","TÓTEMS CENTRALES 1,90 x 3,20 m (2 UND)"),
   sub("STAND PAÍS · TÓTEMS DE FONDO 3,00 x 2,50 m (4 UND)","TÓTEMS DE FONDO 3,00 x 2,50 m (4 UND)"),
   sub("STAND PAÍS · TARIMA Y PISO","TARIMA")]),
("STAND A-01 HIKVISION 4,00 x 7,50 m","STANDS A",_G["STAND A-01 HIKVISION 4,00 x 7,50 m"][1]),
("STAND A-02 GINLONG – SOLIS 4,00 x 7,50 m","STANDS A",_G["STAND A-02 GINLONG – SOLIS 4,00 x 7,50 m"][1]),
("STAND A-03 WEIQIAO (212) 4,00 x 7,50 m","STANDS A",_G["STAND A-03 WEIQIAO (212) 4,00 x 7,50 m"][1]),
("STAND A-04 KONKA 4,00 x 7,50 m","STANDS A",_G["STAND A-04 KONKA 4,00 x 7,50 m"][1]),
("STANDS ESTÁNDAR (40 MÓDULOS, 32 EMPRESAS)","STANDS B",[sub("STAND ESTÁNDAR 9 m² (26 STANDS: 22 UNA CARA ABIERTA + 4 DOS CARAS)"),sub("STAND ESTÁNDAR 18 m² (5 STANDS: 1 UNA CARA + 3 DOS CARAS + 1 TRES CARAS)"),sub("STAND ESTÁNDAR 36 m² BYD (1 STAND)")]),
("PÓRTICO DE INGRESO PRINCIPAL 6,65 x 4,26 m","PÓRTICO DE INGRESO",_G["PÓRTICO DE INGRESO PRINCIPAL 6,65 x 4,26 m"][1]),
("MURAL DECORATIVO (PARED DE INGRESO)","MURAL DECORATIVO",_G["MURAL DECORATIVO (PARED DE INGRESO)"][1]),
("SEÑALÉTICA","SEÑALÉTICA",[sub("BANDEROLAS (6 UND)"),sub("TÓTEMS DIRECCIONALES 0,80 x 1,80 m (4 UND)")]),
("LOGÍSTICA, DESMONTAJE Y ACOMPAÑAMIENTO","GENERAL",_G["LOGÍSTICA, DESMONTAJE Y ACOMPAÑAMIENTO"][1]),
]
def fm(x): return f"{x:.2f}".replace(".",",")
wb=Workbook(); ws=wb.active; ws.title="CTZ-129.2026"
f=lambda **k: Font(name="Arial",size=k.pop("size",9),**k)
th=Side(style="thin",color="A6A6A6"); bx=Border(left=th,right=th,top=th,bottom=th)
dk=PatternFill("solid",fgColor="404040"); grey=PatternFill("solid",fgColor="D9D9D9"); yel=PatternFill("solid",fgColor="FFFF00"); lt=PatternFill("solid",fgColor="F2F2F2")
NUM='#,##0.00;(#,##0.00);"-"'; BL="0000FF"
for c,w in zip("ABCDEFGHIJKLMN",[7,80,7,10,11,13,3,9,9,7,7,11,36,18]): ws.column_dimensions[c].width=w
hdr=[("COTIZACIÓN","CTZ-129.2026"),("FECHA","24/11/2026"),("CLIENTE","FEXPOCRUZ SRL"),("OBRA","PABELLÓN CHINA – EXPOCRUZ 2027"),("UBICACIÓN","PREDIO FERIAL FEXPOCRUZ"),("ESTADO","Borrador")]
for i,(a,b) in enumerate(hdr,1): ws.cell(i,2,f"{a}: {b}").font=f(bold=(i in (1,4)))
ws["B3"].fill=yel
R0=8
H=["N°","ARTEFACTO / PIEZA","UND","CANTIDAD","P.U. (Bs)","SUBTOTAL (Bs)","","LARGO","ALTO","CARAS","N° PZAS","FUENTE","REFERENCIA DE PRECIO","CATEGORÍA"]
for i,h in enumerate(H,1):
    if not h: continue
    c=ws.cell(R0,i,h); c.font=f(bold=True,color="FFFFFF"); c.fill=dk if i<7 else PatternFill("solid",fgColor="7F7F7F"); c.border=bx
    c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
ws.cell(R0-1,8,"Metrado (interno)").font=f(italic=True,color="7F7F7F")
r=R0+1; heads=[]
def style(r,hdr=False):
    for c in range(1,15):
        if c==7: continue
        x=ws.cell(r,c); x.border=bx
        if hdr: x.fill=grey; x.font=f(bold=True)
        x.alignment=Alignment(vertical="top",wrap_text=c in (2,13),horizontal="center" if c in (1,3,10,11,12) else None)
mid=PatternFill("solid",fgColor="EDEDED")
def write_items(items,prefix,cat):
    global r
    k=0; brows=[]; first=r
    for it in items:
        k+=1; code=f"{prefix}.{k}"
        if it[0]=="P":
            _,d,L,A,F,N,u,pu,src,obs=it
            t=d.upper()
            if u=="M2" and L is not None:
                med=f"MEDIDAS: {fm(L)} x {fm(A)} m"+(f", {int(F)} CARAS" if F and F!=1 else "")+(f", {N} PIEZAS" if N!=1 else "")
                t+=f". {med}"
            elif N!=1 and u=="Pza": t+=f". {N} PIEZAS"
            inc=INC.get(src,"INCLUYE MATERIAL, FABRICACIÓN E INSTALACIÓN." if u=="Pza" else "")
            t=(t+". "+inc).replace("..",".")
            if obs: t+=f" ({obs.upper()})"
            for col,v in ((8,L),(9,A),(10,F),(11,N)):
                if v is not None: ws.cell(r,col,v).font=f(color=BL)
            ws.cell(r,4,f"=ROUND(H{r}*I{r}*J{r}*K{r},2)" if u=="M2" else f"=K{r}")
            fu="Estimado" if src==EST else "Histórico"; ref=src
        elif it[0]=="B":
            _,d,L,A,N,m,a=it; u="M2"; pu=PB; fu="Histórico"; ref=S["bas"]+" · volúmenes Drei Haus"
            med=f"ÁREA: {fm(L)} m²" if "área según" in d else f"MEDIDAS: {fm(L)} x {fm(A)} m"+(f", {N} PIEZAS" if N!=1 else "")
            t=f"{d.upper()}: ESTRUCTURA METÁLICA DE BASTIDOR. {med}. INCLUYE FABRICACIÓN E INSTALACIÓN."
            for col,v in ((8,L),(9,A),(10,1),(11,N)): ws.cell(r,col,v).font=f(color=BL)
            ws.cell(r,4,f"=ROUND(H{r}*I{r}*J{r}*K{r},2)"); brows.append((r,m,a))
        elif it[0]=="W":
            _,d,L,A,N,m,a=it; u="M2"; pu=PW; fu="Histórico"; ref=S["dw"]+" · volúmenes Drei Haus"
            med=f"ÁREA: {fm(L)} m²" if "área según" in d else f"MEDIDAS: {fm(L)} x {fm(A)} m"+(f", {N} PIEZAS" if N!=1 else "")
            t=f"{d.upper().replace('BASTIDOR','MURO')}: MURO DE DRYWALL CON ESTRUCTURA DE PERFILES GALVANIZADOS, PLACA DE YESO EN AMBAS CARAS, MASILLADO CON CINTA Y LIJADO, LISTO PARA ADHESIVO. {med}. INCLUYE MATERIAL Y MANO DE OBRA."
            for col,v in ((8,L),(9,A),(10,1),(11,N)): ws.cell(r,col,v).font=f(color=BL)
            ws.cell(r,4,f"=ROUND(H{r}*I{r}*J{r}*K{r},2)"); brows.append((r,m,a))
        elif it[0]=="MDF":
            u="M2"; pu=PM; fu="Histórico"; ref=S["mdf"]
            t="REVESTIMIENTO DE MDF SOBRE LOS BASTIDORES DE ESTE ARTEFACTO. INCLUYE CORTE, FIJACIÓN Y PINTURA BASE."
            ws.cell(r,4,"=ROUND("+"+".join(f"D{rr}" if mm==1 else f"D{rr}*{mm}" for rr,mm,_ in brows if mm)+",2)")
        elif it[0]=="DRY":
            u="M2"; pu=PD; fu="Histórico"; ref=S["dry"]
            t="REVESTIMIENTO DE DRYWALL (PLACA DE YESO) SOBRE LOS BASTIDORES DE ESTE STAND, FIJADO A LA ESTRUCTURA, MASILLADO CON CINTA Y LIJADO, LISTO PARA ADHESIVO. INCLUYE MATERIAL Y MANO DE OBRA."
            ws.cell(r,4,"=ROUND("+"+".join(f"D{rr}" if mm==1 else f"D{rr}*{mm}" for rr,mm,_ in brows if mm)+",2)")
        elif it[0]=="ADH":
            u="M2"; pu=PA; fu="Histórico"; ref=S["adh"]
            t=f"ADHESIVO VINÍLICO SOBRE LOS MUROS DE ESTE STAND{it[1]}. IMPRESIÓN DIGITAL FULL COLOR. INCLUYE COLOCADO."
            ws.cell(r,4,"=ROUND("+"+".join(f"D{rr}" if aa==1 else f"D{rr}*{aa}" for rr,_,aa in brows if aa)+",2)")
        ws.cell(r,1,code); ws.cell(r,2,t); ws.cell(r,3,u.upper()); p=ws.cell(r,5,pu)
        ws.cell(r,6,f"=ROUND(D{r}*E{r},2)"); ws.cell(r,12,fu); ws.cell(r,13,ref)
        ws.cell(r,14,"MOBILIARIO ESPECIAL" if "MOBILIARIO ESPECIAL" in t else cat)
        style(r)
        for c in range(1,15):
            if c==7: continue
            x=ws.cell(r,c)
            if not (c in (8,9,10,11) and x.value is not None): x.font=f(color=BL if c==5 else ("7F7F7F" if c>=12 else None))
        if fu=="Estimado": p.fill=yel
        for c in (4,8,9): ws.cell(r,c).number_format='0.00'
        for c in (5,6): ws.cell(r,c).number_format=NUM
        ws.row_dimensions[r].height=12*(-(-len(t)//82))+4
        r+=1
    return first,r-1
for gi,(title,cat,children) in enumerate(T,1):
    hr=r; ws.cell(r,1,gi); ws.cell(r,2,title); ws.cell(r,14,cat); style(r,True)
    for c_ in range(1,7): ws.cell(r,c_).font=f(bold=True,size=10)
    r+=1
    if children and children[0][0]=="SUB":
        subs=[]
        for si,(_,name,items) in enumerate(children,1):
            sr=r; ws.cell(r,1,f"{gi}.{si}"); ws.cell(r,2,name); ws.cell(r,14,cat); style(r)
            for c_ in range(1,15):
                if c_!=7: ws.cell(r,c_).fill=mid; ws.cell(r,c_).font=f(bold=True)
            r+=1
            a_,b_=write_items(items,f"{gi}.{si}",cat)
            x=ws.cell(sr,6,f"=SUM(F{a_}:F{b_})"); x.number_format=NUM; x.font=f(bold=True)
            subs.append(sr)
        x=ws.cell(hr,6,"="+"+".join(f"F{s_}" for s_ in subs))
    else:
        a_,b_=write_items(children,f"{gi}",cat)
        x=ws.cell(hr,6,f"=SUM(F{a_}:F{b_})")
    x.number_format=NUM; x.font=f(bold=True,size=10)
    heads.append((hr,title,cat)); r+=1
last=r-1
ws.cell(r,5,"TOTAL Bs").font=f(bold=True)
c=ws.cell(r,6,"="+"+".join(f"F{h}" for h,_,_ in heads)); c.font=f(bold=True,size=10); c.number_format=NUM; c.fill=grey; c.border=bx
tot=r; r+=2
ws.cell(r,1,"NOTAS").font=f(bold=True); r+=1
notas=["PRECIOS CON IMPUESTOS DE LEY. LOS PRECIOS UNITARIOS SALEN DE COTIZACIONES FERIALES 2026 (COLUMNA M); LOS MARCADOS EN AMARILLO SON ESTIMADOS.",
"STANDS A Y STAND PAÍS (ESCENARIO Y TÓTEMS): VOLÚMENES SEGÚN PLANILLA DE VOLÚMENES DREI·HAUS, STANDS A EN MURO DE DRYWALL (ESTRUCTURA GALVANIZADA + PLACA, UN SOLO ÍTEM) MÁS ADHESIVO; STAND PAÍS EN BASTIDOR + MDF.",
"INCLUYE LA FABRICACIÓN DEL MOBILIARIO ESPECIAL DISEÑADO POR LOS EXPOSITORES (PLANO 'MUEBLES CHINA'), DENTRO DE CADA STAND.",
"NO INCLUYE MOBILIARIO DE FERIA (MESAS, SILLAS, BANQUETAS, CREDENZAS, COUNTERS ESTÁNDAR, SILLAS DE AUDITORIO, PODIO, CUBOS).",
"NO INCLUYE INSTALACIONES ELÉCTRICAS, ILUMINACIÓN, REFLECTORES, TABLEROS NI PANTALLA LED.",
"NO INCLUYE ALFOMBRA (NI EN STANDS NI EN PASILLOS), PANELERÍA MODULAR (OCTANORM) DE LOS STANDS B NI EL VEHÍCULO DEL STAND A-03.",
"LAS MEDIDAS SIN COTA EN PLANOS SE TOMARON A ESCALA SOBRE LAS ELEVACIONES (INDICADO EN LA LÍNEA); SE AJUSTAN SEGÚN LA MEDICIÓN FINAL.",
"FORMA DE PAGO: 70 % ANTICIPO Y 30 % CONTRA ENTREGA, EN EFECTIVO, CHEQUE O DEPÓSITO. VALIDEZ 5 DÍAS HÁBILES."]
for t in notas:
    c=ws.cell(r,2,t); c.font=f(); c.alignment=Alignment(wrap_text=True); ws.row_dimensions[r].height=24 if len(t)>100 else 13; r+=1
ws.freeze_panes=f"C{R0+1}"
SH="'CTZ-129.2026'"; RNG=f"{SH}!$N${R0+1}:$N${last}"; IMP=f"{SH}!$F${R0+1}:$F${last}"; TXT=f"{SH}!$B${R0+1}:$B${last}"; NR=f"{SH}!$A${R0+1}:$A${last}"
# Resumen por artefacto
s=wb.create_sheet("RESUMEN POR ARTEFACTO")
for col,w in zip("ABCD",[6,62,16,9]): s.column_dimensions[col].width=w
for i,h in enumerate(["N°","ARTEFACTO","SUBTOTAL (Bs)","%"],1):
    x=s.cell(1,i,h); x.font=f(bold=True,color="FFFFFF"); x.fill=dk; x.border=bx
n=len(heads)
for i,(hr,title,cat) in enumerate(heads):
    rr=i+2; s.cell(rr,1,i+1); s.cell(rr,2,title)
    x=s.cell(rr,3,f"={SH}!F{hr}"); x.number_format=NUM; x.font=f(color="008000")
    s.cell(rr,4,f"=IF($C${n+2}=0,0,C{rr}/$C${n+2})").number_format="0.0%"
    for cc in range(1,5):
        s.cell(rr,cc).border=bx
        if cc!=3: s.cell(rr,cc).font=f()
rr=n+2; s.cell(rr,2,"TOTAL").font=f(bold=True); x=s.cell(rr,3,f"=SUM(C2:C{rr-1})"); x.number_format=NUM; x.font=f(bold=True); x.fill=grey
s.cell(rr+1,2,"Control vs. total de la cotización").font=f(italic=True,color="7F7F7F"); x=s.cell(rr+1,3,f"={SH}!F{tot}-C{rr}"); x.number_format=NUM
# Comparativo 2025 (excluye filas de título: N° con punto)
q=wb.create_sheet("COMPARATIVO 2025")
for col,w in zip("ABCDEFG",[26,48,14,16,16,14,40]): q.column_dimensions[col].width=w
q["A1"]="COMPARATIVO CTZ-129.2026 vs. PROPUESTA QUALY 2025 (DDA/AA/7064-2025, Bs 257.643 – desc. 7.643 = 250.000)"; q["A1"].font=f(bold=True,size=11)
hd=["CATEGORÍA","ÍTEMS QUALY 2025","QUALY 2025 (Bs)","CTZ-129 (Bs)","CTZ-129 SIN ALFOMBRA","DIFERENCIA","COMENTARIO"]
for i,h in enumerate(hd,1):
    x=q.cell(3,i,h); x.font=f(bold=True,color="FFFFFF"); x.fill=dk; x.border=bx; x.alignment=Alignment(wrap_text=True,horizontal="center")
comp=[("STANDS B","No incluido en 2025",0,"En 2025 Qualy no cotizó los stands B"),
 ("STAND PAÍS","Arcos 27.725 + Tótems 13.460 + Estructura pantalla 16.410 + Tarima (alquiler) 6.150",63745,"Sin pantalla LED (40.950 en 2025)"),
 ("STANDS A","Custom B1 Solis 20.080 + B2 Hikvision 32.840 + B3 212 18.530 + B4 Konka 23.353",94803,"Qualy B2 con 13 dicroicos; B1 con puerta de depósito"),
 ("PÓRTICO DE INGRESO","Estructura de ingreso 9,10x4,26 (lona + letras luminosas)",19100,""),
 ("MURAL DECORATIVO","Ingreso China (frontal) 9,70x3,50 con 7 reflectores",24545,""),
 ("SEÑALÉTICA","Tótem 4 x 1.600 + Banderolas 6 x 1.350",14500,""),
 ("GENERAL","Incluido en precios",0,"Logística, desmontaje y técnico"),
 ("MOBILIARIO ESPECIAL","No incluido",0,"Counter y vitrina Hikvision, display y atriles Solis")]
rr=4
for cat,txt,qv,com in comp:
    q.cell(rr,1,cat); q.cell(rr,2,txt); q.cell(rr,3,qv)
    q.cell(rr,4,f'=SUMIFS({IMP},{RNG},"{cat}",{SH}!$L${R0+1}:$L${last},"?*")')
    q.cell(rr,5,f'=D{rr}-SUMIFS({IMP},{RNG},"{cat}",{SH}!$L${R0+1}:$L${last},"?*",{TXT},"ALFOMBRA*")'); q.cell(rr,6,f"=E{rr}-C{rr}"); q.cell(rr,7,com)
    for cc in range(1,8):
        x=q.cell(rr,cc); x.border=bx; x.font=f(color=BL if cc==3 else None); x.alignment=Alignment(wrap_text=cc in (2,7),vertical="top")
    for cc in (3,4,5,6): q.cell(rr,cc).number_format=NUM
    rr+=1
q.cell(rr,1,"TOTAL COMPARABLE").font=f(bold=True)
for cc in (3,4,5,6):
    x=q.cell(rr,cc,f"=SUM({chr(64+cc)}4:{chr(64+cc)}{rr-1})"); x.number_format=NUM; x.font=f(bold=True); x.fill=grey; x.border=bx
q.cell(rr+1,4,f"={SH}!F{tot}-D{rr}").number_format=NUM; q.cell(rr+1,1,"Control vs. total").font=f(italic=True,color="7F7F7F")
rr+=3
for t in ["Fuente Qualy: Cotización N° 0-134 del 31/07/2025 y propuesta DDA/AA/7064-2025 (precios con factura).",
          "Pantalla LED alquiler 2025: Bs 40.950 – excluida de CTZ-129.","Anotación a mano en la carátula: 112 / 250 / 138."]:
    q.cell(rr,1,t).font=f(italic=True,color="7F7F7F"); rr+=1
wb.save(OUT); print(OUT)
