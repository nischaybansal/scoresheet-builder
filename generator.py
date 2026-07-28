"""
Scoresheet generator engine.
Given a competency framework (list of (competency, behaviour)) and a list of
selected tools, builds the multi-tab assessor scoresheet workbook:
  - one sheet per selected ONLINE tool (OPQ/MQ/SJT) with the paste-stens engine
  - one sheet per selected OFFLINE tool (Case Study/Role Play/BEI) as a blank scaffold
  - a Detailed Integration sheet wiring every tool score into place
Returns an openpyxl Workbook.
"""
from itertools import groupby
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

# ---- constant SHL dimension libraries (captured from the sample workbook) ----
OPQ_DIMS = ["Persuasive","Controlling","Outspoken","Independent Minded","Outgoing","Affiliative",
"Socially Confident","Modest","Democratic","Caring","Data Rational","Evaluative","Behavioural",
"Conventional","Conceptual","Innovative","Variety Seeking","Adaptable","Forward Thinking",
"Detail Conscious","Conscientious","Rule Following","Relaxed","Worrying","Tough Minded","Optimistic",
"Trusting","Emotionally Controlled","Vigorous","Competitive","Achieving","Decisive"]
MQ_DIMS = ["Level of Activity","Achievement","Competition","Fear of Failure","Power","Immersion",
"Commercial Outlook","Affiliation","Recognition","Personal Principles","Ease and Security",
"Personal Growth","Interest","Flexibility","Autonomy","Material Reward","Progression","Status"]
SJT_VARIANTS = {
    "Managerial": {"group":"Scenarios Management",
        "dims":["Managerial Judgement","Managing Objectives","People Management","Reputation Management",
                "Big Picture","Delegative","One To One","Team","Personal Recognition","Company Protocol"]},
    "Executive": {"group":"Scenarios Executive Profile",
        "dims":["Corporate Management","Managerial Judgement","Managing Objectives","People Management"]},
}
DEFAULT_SJT_VARIANT = "Managerial"

# ---- tool registry: canonical order = online first, then offline ----
# kind: 'online' (paste-stens), 'bars' (rubric anchors), 'bei' (interview questions)
# score_col = the column on the tool sheet holding the behaviour's 1-5 rating that DI reads
VERIFY_DIMS = ["Inductive Score","Deductive Score","Numerical Score"]

# ---- Onlines Data sheet layout (mirrors the SHL export template) ----
OD_SHEET = "Onlines Data"
OD_FIRST_DATA_ROW = 4          # participants pasted from row 4 down
OD_LAST_DATA_ROW  = 503        # 500 participants headroom (dynamic - blanks ignored)
OD_TRAIT_FIRST_COL = "F"       # trait columns start at F
OD_TRAIT_LAST_COL  = "CZ"      # headroom for extra export columns
EMAIL_REF = "'Detailed Integration'!$B$2"   # selected participant email
OD_RAW_HEADERS = {             # row-2 raw export headers where known (cosmetic only)
"Level of Activity":"Level of Activity (E1)","Achievement":"Achievement (E2)","Competition":"Competition (E3)",
"Fear of Failure":"Fear of Failure (E4)","Power":"Power (E5)","Immersion":"Immersion (E6)",
"Commercial Outlook":"Commercial Outlook (E7)","Affiliation":"Affiliation (S1)","Recognition":"Recognition (S2)",
"Personal Principles":"Personal Principles (S3)","Ease and Security":"Ease and Security (S4)",
"Personal Growth":"Personal Growth (S5)","Interest":"Interest (I1)","Flexibility":"Flexibility (I2)",
"Autonomy":"Autonomy (I3)","Material Reward":"Material Reward (X1)","Progression":"Progression (X2)",
"Status":"Status (X3)","Consistency Measure":"OPQ32 - Consistency Measure (CNS)"}

TOOLS = {
    "OPQ":              {"kind":"online","dims":OPQ_DIMS,   "family":"online"},
    "MQ":               {"kind":"online","dims":MQ_DIMS,    "family":"online"},
    "SJT":              {"kind":"online","dims":None,       "family":"online"},  # dims from SJT_VARIANTS
    "Verify":           {"kind":"online","dims":VERIFY_DIMS,"family":"online"},
    "Case Study":       {"kind":"bars","score_col":"H","family":"offline"},
    "Group Discussion": {"kind":"bars","score_col":"H","family":"offline"},
    "Inbox Simulation": {"kind":"bars","score_col":"H","family":"offline"},
    "Written Analysis": {"kind":"bars","score_col":"H","family":"offline"},
    "Role Play":        {"kind":"bars","score_col":"H","family":"offline"},
    "BEI":              {"kind":"bei","score_col":"F","family":"offline"},
}
CANONICAL = ["OPQ","MQ","SJT","Verify",
             "Case Study","Group Discussion","Inbox Simulation","Written Analysis","Role Play","BEI"]
def _avg_col(name):
    """Competency-average column on an online sheet (OPQ has an extra Direction column)."""
    return "L" if name=="OPQ" else "K"

HR = 6          # header row on every sheet
FB = HR + 1     # first behaviour row (behaviour i -> row FB+i on every tool sheet)

# ---- styling ----
ARIAL="Montserrat"   # all cells use Montserrat 9 (must be installed for Excel to render it)
C_HDR=PatternFill("solid",fgColor="1F3864"); C_SUB=PatternFill("solid",fgColor="D9E1F2")
C_IN=PatternFill("solid",fgColor="FFF2CC");  C_LOCK=PatternFill("solid",fgColor="E2EFDA")
C_LEG=PatternFill("solid",fgColor="FCE4D6"); C_PASTE=PatternFill("solid",fgColor="DDEBF7")
C_DIMBG=PatternFill("solid",fgColor="F2F2F2"); C_BLACK=PatternFill("solid",fgColor="000000")
F_HDR=Font(name=ARIAL,size=9,bold=True,color="FFFFFF"); F_B=Font(name=ARIAL,size=9,bold=True)
F_N=Font(name=ARIAL,size=9); F_TITLE=Font(name=ARIAL,size=9,bold=True,color="1F3864")
F_IT=Font(name=ARIAL,size=9,italic=True); F_TAG=Font(name=ARIAL,size=9,bold=True,color="1F3864")
_thin=Side(style="thin",color="BFBFBF"); BD=Border(_thin,_thin,_thin,_thin)
A_WRAP=Alignment(wrap_text=True,vertical="top"); A_CTR=Alignment(horizontal="center",vertical="center")
A_WC=Alignment(wrap_text=True,vertical="center",horizontal="center")
A_WV=Alignment(wrap_text=True,vertical="center")

def _ref(sheet):
    return f"'{sheet}'" if " " in sheet else sheet

def _hdr(ws,row,cols):
    for c in cols:
        x=ws[f"{c}{row}"]; x.font=F_HDR; x.fill=C_HDR; x.alignment=A_WC; x.border=BD

def _title(ws,text,legend,span):
    ws["A1"]=text; ws["A1"].font=F_TITLE
    ws["A3"]=legend; ws["A3"].font=F_N; ws["A3"].fill=C_LEG
    ws.merge_cells(f"A3:{span}3"); ws["A3"].alignment=A_WV

def _cellsetup(ws,row,cols):
    for c in cols:
        ws[f"{c}{row}"].border=BD; ws[f"{c}{row}"].font=F_N; ws[f"{c}{row}"].alignment=A_WRAP

# ---------- ONLINE tool sheet (paste-stens engine) ----------
TRAIT_SLOTS = 8   # max mapped traits per competency on online tools

def _blocks(framework):
    """[(comp, [behaviours])] preserving order (framework arrives clustered)."""
    out=[]
    for comp,beh in framework:
        if out and out[-1][0]==comp: out[-1][1].append(beh)
        else: out.append((comp,[beh]))
    return out

def _block_geometry(framework):
    """Row layout shared by all online sheets: block height = max(n_behaviours, TRAIT_SLOTS).
    Returns [(comp, behaviours, start_row, end_row)]."""
    geo=[]; r=FB
    for comp,behs in _blocks(framework):
        h=max(len(behs),TRAIT_SLOTS)
        geo.append((comp,behs,r,r+h-1)); r+=h
    return geo

def _build_online(wb,name,dims,framework):
    from openpyxl.styles import Protection
    has_dir = (name=="OPQ")
    ws=wb.create_sheet(name)
    disp=name
    key=("SCORING KEY (constant): Positive trait -> Score = roundup(Sten/2): 1-2=1 · 3-4=2 · 5-6=3 · 7-8=4 · 9-10=5.  "
         "Negative trait -> inverse: 1-2=5 · 3-4=4 · 5-6=3 · 7-8=2 · 9-10=1.") if has_dir else \
        ("SCORING KEY (constant): Score = roundup(Sten/2): 1-2=1 · 3-4=2 · 5-6=3 · 7-8=4 · 9-10=5.")
    avg=_avg_col(name)
    if name=="SJT":
        # label which variant this sheet was generated for
        disp=f"SJT ({[k for k,v in SJT_VARIANTS.items() if v['dims']==dims][0]})" if any(v['dims']==dims for v in SJT_VARIANTS.values()) else name
    _title(ws,f"{disp}  —  online psychometric",key,avg)
    ws["A4"]=("Scores flow automatically from the 'Onlines Data' sheet for the participant selected in "
              "Detailed Integration. Consultant maps up to "
              f"{TRAIT_SLOTS} traits per competency and writes the Rating 1-5 rubric descriptions"
              +(" and sets Direction (Positive by default)" if has_dir else "")
              +". This sheet is protected - only mapping"
              +(", direction" if has_dir else "")+" and rubric cells are editable.")
    ws["A4"].font=F_IT; ws.merge_cells(f"A4:{avg}4"); ws["A4"].alignment=A_WV
    heads=["Competency","Behaviour",f"Mapped {name} Trait"]+(["Direction"] if has_dir else [])+ \
          ["Rating 1","Rating 2","Rating 3","Rating 4","Rating 5","Sten (auto)","Score (1-5)","Competency Avg"]
    cols=[openpyxl.utils.get_column_letter(i+1) for i in range(len(heads))]
    dir_col="D" if has_dir else None
    rub_cols=cols[4:9] if has_dir else cols[3:8]
    sten_col=cols[-3]; score_col=cols[-2]
    for c,h in zip(cols,heads): ws[f"{c}{HR}"]=h
    _hdr(ws,HR,cols)
    # dimension list (dropdown source) + the selected participant's stens beside it
    dim_col=openpyxl.utils.get_column_letter(len(cols)+2)
    psten_col=openpyxl.utils.get_column_letter(len(cols)+3)
    od_data=f"'{OD_SHEET}'!${OD_TRAIT_FIRST_COL}${OD_FIRST_DATA_ROW}:${OD_TRAIT_LAST_COL}${OD_LAST_DATA_ROW}"
    od_names=f"'{OD_SHEET}'!${OD_TRAIT_FIRST_COL}$3:${OD_TRAIT_LAST_COL}$3"
    od_email=f"'{OD_SHEET}'!$D${OD_FIRST_DATA_ROW}:$D${OD_LAST_DATA_ROW}"
    ws[f"{dim_col}{HR}"]=f"{name} Dimensions"; ws[f"{psten_col}{HR}"]="Sten (auto)"
    _hdr(ws,HR,[dim_col,psten_col])
    for i,dim in enumerate(dims):
        r=FB+i
        ws[f"{dim_col}{r}"]=dim; ws[f"{dim_col}{r}"].font=F_N; ws[f"{dim_col}{r}"].border=BD; ws[f"{dim_col}{r}"].fill=C_DIMBG
        p=ws[f"{psten_col}{r}"]
        _ix=(f'INDEX({od_data},MATCH({EMAIL_REF},{od_email},0),MATCH({dim_col}{r},{od_names},0))')
        p.value=f'=IF({EMAIL_REF}="","",IFERROR(IF({_ix}="","",{_ix}),""))'
        p.font=F_N; p.border=BD; p.fill=C_LOCK; p.alignment=A_CTR
    dim_last=FB+len(dims)-1
    unlocked=Protection(locked=False)
    geo=_block_geometry(framework)
    for comp,behs,gs,ge in geo:
        for k in range(ge-gs+1):
            r=gs+k
            if k<len(behs):
                ws[f"A{r}"]=comp; ws[f"B{r}"]=behs[k]
            if k<TRAIT_SLOTS:
                ws[f"C{r}"]=""; ws[f"C{r}"].fill=C_IN
                if has_dir:
                    ws[f"{dir_col}{r}"]="Positive"; ws[f"{dir_col}{r}"].fill=C_IN
                for rc in rub_cols:
                    ws[f"{rc}{r}"]=""; ws[f"{rc}{r}"].fill=C_SUB
                _ix=(f'INDEX({od_data},MATCH({EMAIL_REF},{od_email},0),MATCH(C{r},{od_names},0))')
                ws[f"{sten_col}{r}"]=(f'=IF(OR({EMAIL_REF}="",C{r}=""),"",'
                    f'IFERROR(IF({_ix}="","",{_ix}),""))')
                ws[f"{sten_col}{r}"].fill=C_LOCK
                if has_dir:
                    ws[f"{score_col}{r}"]=(f'=IF({sten_col}{r}="","",IF({dir_col}{r}="Negative",'
                        f'6-ROUNDUP({sten_col}{r}/2,0),ROUNDUP({sten_col}{r}/2,0)))')
                else:
                    ws[f"{score_col}{r}"]=f'=IF({sten_col}{r}="","",ROUNDUP({sten_col}{r}/2,0))'
                ws[f"{score_col}{r}"].fill=C_LOCK
                for cc in (["C"]+([dir_col] if has_dir else [])+rub_cols):
                    ws[f"{cc}{r}"].protection=unlocked
            _cellsetup(ws,r,cols)
            ctr_cols=[sten_col,score_col]+([dir_col] if has_dir else [])
            for cc in ctr_cols: ws[f"{cc}{r}"].alignment=A_CTR
        g=ws[f"{avg}{gs}"]
        g.value=f'=IFERROR(ROUND(AVERAGE({score_col}{gs}:{score_col}{ge}),0),"")'
        g.fill=C_LOCK; g.font=F_B; g.alignment=A_CTR
        if ge>gs: ws.merge_cells(f"{avg}{gs}:{avg}{ge}")
        for rr in range(gs,ge+1): ws[f"{avg}{rr}"].border=BD
    last=geo[-1][3]
    dv=DataValidation(type="list",formula1=f"=${dim_col}${FB}:${dim_col}${dim_last}",allow_blank=True)
    ws.add_data_validation(dv); dv.add(f"C{FB}:C{last}")
    if has_dir:
        dv2=DataValidation(type="list",formula1='"Positive,Negative"',allow_blank=True)
        ws.add_data_validation(dv2); dv2.add(f"{dir_col}{FB}:{dir_col}{last}")
    widths=[20,38,22]+([10] if has_dir else [])+[17,17,17,17,17,10,9,12]
    for c,w in zip(cols,widths): ws.column_dimensions[c].width=w
    ws.column_dimensions[dim_col].width=24
    ws.column_dimensions[psten_col].width=11
    ws.freeze_panes=f"A{FB}"
    ws.protection.sheet=True     # lock the sheet; only unlocked cells editable

# ---------- OFFLINE: BARS rubric sheet (Case Study / Role Play) ----------
def _build_bars(wb,name,framework):
    ws=wb.create_sheet(name)
    _title(ws,f"{name}  —  scoring guidelines",
        "Consultant writes the anchors under Rating 1-5. Assessor enters observed Rating (1-5) + comments.","I")
    cols=["A","B","C","D","E","F","G","H","I"]
    heads=["Competency","Behaviour","Rating 1","Rating 2","Rating 3","Rating 4","Rating 5","Rating","Comments"]
    for c,h in zip(cols,heads): ws[f"{c}{HR}"]=h
    _hdr(ws,HR,cols)
    for i,(comp,beh) in enumerate(framework):
        r=FB+i; ws[f"A{r}"]=comp; ws[f"B{r}"]=beh
        for c in ["C","D","E","F","G"]: ws[f"{c}{r}"]=""; ws[f"{c}{r}"].fill=C_SUB
        ws[f"H{r}"]=""; ws[f"H{r}"].fill=C_IN; ws[f"I{r}"]=""; ws[f"I{r}"].fill=C_IN
        _cellsetup(ws,r,cols); ws[f"H{r}"].alignment=A_CTR
    for c,w in zip(cols,[18,36,18,18,18,18,18,8,28]): ws.column_dimensions[c].width=w
    ws.freeze_panes=f"A{FB}"

# ---------- OFFLINE: BEI sheet ----------
def _build_bei(wb,name,framework):
    ws=wb.create_sheet(name)
    _title(ws,f"{name}  —  Behavioural Event Interview",
        "Consultant writes 1-3 BEI questions per behaviour. Assessor enters Rating (1-5) + comments.","G")
    cols=["A","B","C","D","E","F","G"]
    heads=["Competency","Behaviour","BEI Question 1","BEI Question 2","BEI Question 3","Rating","Comments"]
    for c,h in zip(cols,heads): ws[f"{c}{HR}"]=h
    _hdr(ws,HR,cols)
    for i,(comp,beh) in enumerate(framework):
        r=FB+i; ws[f"A{r}"]=comp; ws[f"B{r}"]=beh
        for c in ["C","D","E"]: ws[f"{c}{r}"]=""; ws[f"{c}{r}"].fill=C_SUB
        ws[f"F{r}"]=""; ws[f"F{r}"].fill=C_IN; ws[f"G{r}"]=""; ws[f"G{r}"].fill=C_IN
        _cellsetup(ws,r,cols); ws[f"F{r}"].alignment=A_CTR
    for c,w in zip(cols,[18,36,26,26,26,8,28]): ws.column_dimensions[c].width=w
    ws.freeze_panes=f"A{FB}"

# ---------- Onlines Data (assessor pastes the export here) ----------
def _build_onlines_data(wb,sjt_variant=DEFAULT_SJT_VARIANT):
    ws=wb.create_sheet(OD_SHEET)
    sv=SJT_VARIANTS[sjt_variant]
    sjt_raw={d:f"{d}-STEN" for d in sv["dims"]}
    OD_RAW_HEADERS.update(sjt_raw)
    groups=[("MQ.M5 Profile",MQ_DIMS),
            ("OPQ32 Profile",OPQ_DIMS+["Consistency Measure"]),
            (sv["group"],sv["dims"]),
            ("Verify",["Inductive Percentile","Inductive Score","Deductive percentile",
                       "Deductive Score","Numerical Percentile","Numerical Score","Overall"])]
    ident=["First Name","Last Name","Name","Email","Status"]
    for i,h in enumerate(ident):
        col=openpyxl.utils.get_column_letter(i+1)
        ws.merge_cells(f"{col}1:{col}3")
        c=ws[f"{col}1"]; c.value=h; c.font=F_HDR; c.fill=C_HDR; c.alignment=A_WC
        for r in (1,2,3): ws[f"{col}{r}"].border=BD
    col_idx=6
    for gname,dims in groups:
        gstart=col_idx
        for d in dims:
            col=openpyxl.utils.get_column_letter(col_idx)
            ws[f"{col}2"]=OD_RAW_HEADERS.get(d,d)          # raw export header (cosmetic)
            ws[f"{col}3"]=d                                 # clean trait name (lookup key)
            ws[f"{col}2"].font=F_IT; ws[f"{col}3"].font=F_B
            for r in (2,3):
                ws[f"{col}{r}"].border=BD; ws[f"{col}{r}"].alignment=A_WV
            ws[f"{col}3"].fill=C_DIMBG
            ws.column_dimensions[col].width=16
            col_idx+=1
        gend=col_idx-1
        c1=openpyxl.utils.get_column_letter(gstart); c2=openpyxl.utils.get_column_letter(gend)
        if gend>gstart: ws.merge_cells(f"{c1}1:{c2}1")
        h=ws[f"{c1}1"]; h.value=gname; h.font=F_HDR; h.fill=C_HDR; h.alignment=A_WC
        for cc in range(gstart,gend+1):
            ws[f"{openpyxl.utils.get_column_letter(cc)}1"].border=BD
    note=ws[f"A{OD_FIRST_DATA_ROW+0}"]
    ws.freeze_panes=f"F{OD_FIRST_DATA_ROW}"
    # light input tint on the first visible data rows so assessors know where to paste
    for r in range(OD_FIRST_DATA_ROW,OD_FIRST_DATA_ROW+3):
        for cidx in range(1,col_idx):
            ws[f"{openpyxl.utils.get_column_letter(cidx)}{r}"].fill=C_PASTE
    for c,w in zip(["A","B","C","D","E"],[14,14,22,30,12]): ws.column_dimensions[c].width=w
    return ws

# ---------- T Chart ----------
TCHART_GUIDELINES=("Guidelines:\n"
"1. To be written in Second Person 'He/She'.\n"
"2. Mention strengths of the person (avoid replication of the behavioral indicators or OPQ statements)\n"
"3. Highlight the area of development such that it is (a) specific (b) actionable (c) constructive such that "
"it does not sound prescriptive (avoid 'should') and inspires the participant\n"
"4. Record critical evidences for these areas to justify being strength/areas of development and write them below.\n"
"5. For each development area, write two actionable development tips.\n"
"6. Appropriate linkage to the current or future role (if necessary)\n"
"7. Correctness/ accuracy (grammar, sentence construction, spell check)\n"
"8. 1 Strength area atleast (only if Rating 3-5) and 1 Development Area at least")

def _build_tchart(wb,framework):
    ws=wb.create_sheet("T Chart")
    comps=[]
    for c,_ in framework:
        if c not in comps: comps.append(c)
    # guidelines block
    ws.merge_cells("B2:D12")
    g=ws["B2"]; g.value=TCHART_GUIDELINES; g.font=F_N; g.fill=C_LEG
    g.alignment=Alignment(wrap_text=True,vertical="top")
    for row in ws["B2:D12"]:
        for c in row: c.border=BD
    # summary block
    ws.merge_cells("B15:D15"); t=ws["B15"]; t.value="T Chart Summary"; t.font=F_HDR; t.fill=C_HDR; t.alignment=A_WC
    ws.merge_cells("B16:D16"); s=ws["B16"]; s.value="Strength Areas ( Min. 3 rating)"; s.font=F_B; s.fill=C_SUB; s.alignment=A_WV
    ws.merge_cells("B23:D23"); d=ws["B23"]; d.value="Development Areas (at least 1 competency)"; d.font=F_B; d.fill=C_SUB; d.alignment=A_WV
    rows=[(17,"Strength Area 1",True),(18,"Evidence 1",False),(19,"Evidence 2",False),
          (20,"Strength Area 2",True),(21,"Evidence 1",False),(22,"Evidence 2",False),
          (24,"Development Area 1",True),(25,"Evidence 1",False),(26,"Evidence 2",False),
          (27,"Development Area 2",True),(28,"Evidence 1",False),(29,"Evidence 2",False)]
    for r,label,is_dd in rows:
        ws[f"B{r}"]=label; ws[f"B{r}"].font=F_B; ws[f"B{r}"].border=BD
        ws.merge_cells(f"C{r}:D{r}")
        cell=ws[f"C{r}"]; cell.value=""; cell.fill=C_IN
        cell.alignment=Alignment(wrap_text=True,vertical="top")
        for cc in ["C","D"]: ws[f"{cc}{r}"].border=BD
        if not is_dd:
            ws.row_dimensions[r].height=95
    # competency list (hidden helper column) + dropdowns on the four area rows
    for i,comp in enumerate(comps):
        ws[f"H{FB+i}"]=comp
    ws.column_dimensions["H"].hidden=True
    dv=DataValidation(type="list",formula1=f"=$H${FB}:$H${FB+len(comps)-1}",allow_blank=True)
    ws.add_data_validation(dv)
    for r in (17,20,24,27): dv.add(f"C{r}")
    for c,w in zip(["A","B","C","D"],[6,24,38,60]): ws.column_dimensions[c].width=w
    return ws

# ---------- Detailed Integration ----------
def _build_di(wb,selected,framework,tool_comps):
    di=wb.create_sheet("Detailed Integration")
    di["A1"]="DETAILED INTEGRATION"; di["A1"].font=F_TITLE
    od_email=f"'{OD_SHEET}'!$D${OD_FIRST_DATA_ROW}:$D${OD_LAST_DATA_ROW}"
    od_name=f"'{OD_SHEET}'!$C${OD_FIRST_DATA_ROW}:$C${OD_LAST_DATA_ROW}"
    tool_cols=[t for t in CANONICAL if t in selected]   # online-first ordering
    has_online=any(TOOLS[t]["family"]=="online" for t in tool_cols)
    # participant block: email dropdown -> name auto; assessor + date manual
    for lab,r in [("Participant Email",2),("Participant Name",3),("Assessor Name",4),("Date of Scoring",5)]:
        di[f"A{r}"]=lab; di[f"A{r}"].font=F_B; di[f"B{r}"]=""; di[f"B{r}"].border=BD
    if has_online:
        di["B2"].fill=C_IN
        dv_email=DataValidation(type="list",formula1=f"={od_email}",allow_blank=True)
        di.add_data_validation(dv_email); dv_email.add("B2")
        di["B3"]=f'=IFERROR(INDEX({od_name},MATCH($B$2,{od_email},0)),"")'
        di["B3"].fill=C_LOCK
    else:
        di["B2"].fill=C_IN; di["B3"].fill=C_IN
    di["B4"].fill=C_IN; di["B5"].fill=C_IN
    H=7                                   # header row (labels use 2-5, tags row 6)
    fixed=["Competency","Behaviour","Behaviour Rating","Competency Rating"]
    heads=fixed+tool_cols
    cols=[openpyxl.utils.get_column_letter(i+1) for i in range(len(heads))]
    for c,h in zip(cols,heads): di[f"{c}{H}"]=h
    _hdr(di,H,cols)
    # Consistency Measure (OPQ validity indicator) directly above the OPQ heading
    if "OPQ" in tool_cols:
        oc=cols[4+tool_cols.index("OPQ")]
        od_data=f"'{OD_SHEET}'!${OD_TRAIT_FIRST_COL}${OD_FIRST_DATA_ROW}:${OD_TRAIT_LAST_COL}${OD_LAST_DATA_ROW}"
        od_names=f"'{OD_SHEET}'!${OD_TRAIT_FIRST_COL}$3:${OD_TRAIT_LAST_COL}$3"
        cell=di[f"{oc}{H-1}"]
        cell.value=(f'=IF($B$2="","","Consistency: "&IFERROR(INDEX({od_data},'
                    f'MATCH($B$2,{od_email},0),MATCH("Consistency Measure",{od_names},0)),"-"))')
        cell.font=F_TAG; cell.alignment=A_CTR
    start=H+1; di_row=start; groups=[]
    beh_col=cols[2]; comp_rating_col=cols[3]
    # per-tool row maps built from each tool's FILTERED framework
    tool_rows={}
    for t in tool_cols:
        filt=[(c,b) for c,b in framework if c in tool_comps[t]]
        if TOOLS[t]["family"]=="online":
            tool_rows[t]={comp:gs for comp,behs,gs,ge in _block_geometry(filt)} if filt else {}
        else:
            tool_rows[t]={(c,b):FB+i for i,(c,b) in enumerate(filt)}
    def _black(dcol,gs,ge):
        for rr in range(gs,ge+1):
            cell=di[f"{dcol}{rr}"]; cell.fill=C_BLACK; cell.border=BD
        if ge>gs: di.merge_cells(f"{dcol}{gs}:{dcol}{ge}")
    for comp,items in groupby(framework,key=lambda x:x[0]):
        items=list(items); gs=di_row
        for (c,beh) in items:
            di[f"{cols[0]}{di_row}"]=comp; di[f"{cols[1]}{di_row}"]=beh
            di[f"{beh_col}{di_row}"]=""; di[f"{beh_col}{di_row}"].fill=C_IN
            for k,t in enumerate(tool_cols):
                if TOOLS[t]["family"]=="online" or comp not in tool_comps[t]:
                    continue  # online + unmapped handled at group level
                dcol=cols[4+k]; sc=TOOLS[t]["score_col"]; trow=tool_rows[t][(comp,beh)]
                di[f"{dcol}{di_row}"]=f'=IF({_ref(t)}!{sc}{trow}="","",{_ref(t)}!{sc}{trow})'
                di[f"{dcol}{di_row}"].fill=C_LOCK
            _cellsetup(di,di_row,cols)
            for cc in cols[2:]: di[f"{cc}{di_row}"].alignment=A_CTR
            di_row+=1
        ge=di_row-1
        for k,t in enumerate(tool_cols):
            dcol=cols[4+k]
            if comp not in tool_comps[t]:
                _black(dcol,gs,ge)                # not mapped -> one black block
                continue
            if TOOLS[t]["family"]!="online":
                continue
            blk_start=tool_rows[t][comp]
            cell=di[f"{dcol}{gs}"]
            cell.value=f'=IF({_ref(t)}!{_avg_col(t)}{blk_start}="","",{_ref(t)}!{_avg_col(t)}{blk_start})'
            cell.fill=C_LOCK; cell.font=F_B; cell.alignment=A_CTR
            if ge>gs: di.merge_cells(f"{dcol}{gs}:{dcol}{ge}")
            for rr in range(gs,ge+1): di[f"{dcol}{rr}"].border=BD
        g=di[f"{comp_rating_col}{gs}"]
        g.value=f'=IFERROR(ROUND(AVERAGE({beh_col}{gs}:{beh_col}{ge}),0),"")'
        g.fill=C_LOCK; g.font=F_B; g.alignment=A_CTR
        if ge>gs: di.merge_cells(f"{comp_rating_col}{gs}:{comp_rating_col}{ge}")
        for rr in range(gs,ge+1): di[f"{comp_rating_col}{rr}"].border=BD
        groups.append((comp,gs,ge)); di_row=ge+1
    widths=[20,46,15,16]+[10]*len(tool_cols)
    for c,w in zip(cols,widths): di.column_dimensions[c].width=w
    di.freeze_panes=f"A{start}"
    # summary table
    sr=di_row+2; di[f"A{sr}"]="COMPETENCY SUMMARY"; di[f"A{sr}"].font=F_B
    shd=sr+1
    for c,h in zip(["A","B","C","D","E"],["Competency","Individual","Cohort","Ideal","Previous DC"]): di[f"{c}{shd}"]=h
    _hdr(di,shd,["A","B","C","D","E"])
    for k,(comp,gs,ge) in enumerate(groups):
        rr=shd+1+k; di[f"A{rr}"]=comp
        di[f"B{rr}"]=f'=IF({comp_rating_col}{gs}="","",{comp_rating_col}{gs})'; di[f"B{rr}"].fill=C_LOCK
        for c in ["C","D","E"]: di[f"{c}{rr}"]=""; di[f"{c}{rr}"].fill=C_IN
        for c in ["A","B","C","D","E"]:
            di[f"{c}{rr}"].border=BD; di[f"{c}{rr}"].font=F_N
            if c!="A": di[f"{c}{rr}"].alignment=A_CTR

# ---------- public API ----------
def _cluster_by_competency(framework):
    """Group all behaviours under the same competency together (first-appearance order),
    so the Detailed Integration competency average is always taken per competency, even
    if the uploaded framework interleaves competencies."""
    order, groups = [], {}
    for comp, beh in framework:
        if comp not in groups:
            groups[comp] = []; order.append(comp)
        groups[comp].append(beh)
    return [(c, b) for c in order for b in groups[c]]

def build_workbook(framework, selected_tools, sjt_variant=DEFAULT_SJT_VARIANT, tool_comps=None):
    """framework: list of (competency, behaviour). selected_tools: subset of CANONICAL.
    sjt_variant: "Managerial" or "Executive" (shapes the SJT sheet + Onlines Data columns).
    tool_comps: optional {tool: [competencies]} - a tool sheet only carries the competencies
    mapped to it; unmapped competencies show as a black block in Detailed Integration.
    Tools not in the dict get every competency."""
    if sjt_variant not in SJT_VARIANTS: sjt_variant=DEFAULT_SJT_VARIANT
    selected=[t for t in CANONICAL if t in selected_tools]
    framework=_cluster_by_competency(framework)   # ensures per-competency averaging
    all_comps=[]
    for c,_ in framework:
        if c not in all_comps: all_comps.append(c)
    tc={}
    for t in selected:
        chosen=(tool_comps or {}).get(t)
        tc[t]=set(chosen) if chosen is not None else set(all_comps)
        if not any(c in tc[t] for c in all_comps):
            raise ValueError(f"'{t}' has no competencies mapped - untick the tool instead, "
                             f"or map at least one competency to it.")
    wb=openpyxl.Workbook(); wb.remove(wb.active)
    has_online=any(TOOLS[t]["family"]=="online" for t in selected)
    if has_online:
        _build_onlines_data(wb,sjt_variant)  # assessor pastes the export here (first sheet)
    for name in selected:
        spec=TOOLS[name]
        filt=[(c,b) for c,b in framework if c in tc[name]]
        if spec["kind"]=="online":
            dims=SJT_VARIANTS[sjt_variant]["dims"] if name=="SJT" else spec["dims"]
            _build_online(wb,name,dims,filt)
        elif spec["kind"]=="bars": _build_bars(wb,name,filt)
        elif spec["kind"]=="bei":  _build_bei(wb,name,filt)
    _build_tchart(wb,framework)
    _build_di(wb,selected,framework,tc)
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines=False          # no gridlines
    return wb

def build_template():
    """Blank 2-column framework template workbook."""
    wb=openpyxl.Workbook(); ws=wb.active; ws.title="Competency Framework"
    ws.sheet_view.showGridLines=False
    ws["A1"]="COMPETENCY FRAMEWORK — upload template"; ws["A1"].font=F_TITLE
    ws["A2"]=("One row per behaviour. Repeat the competency name on each of its behaviour rows. "
              "Add as many rows as you need.")
    ws["A2"].fill=C_LEG; ws["A2"].font=F_IT; ws.merge_cells("A2:B2"); ws["A2"].alignment=A_WV
    ws["A4"]="Competency"; ws["B4"]="Behaviour"
    for c in ["A4","B4"]: ws[c].font=F_HDR; ws[c].fill=C_HDR; ws[c].alignment=Alignment(vertical="center"); ws[c].border=BD
    ex=[("Drives Results","Sets clear goals and holds the team accountable for outcomes."),
        ("Drives Results","Removes obstacles and reprioritises to keep work moving toward targets."),
        ("Builds Collaboration","Works across teams to solve shared problems and align on decisions.")]
    for i,(c,b) in enumerate(ex):
        r=5+i; ws[f"A{r}"]=c; ws[f"B{r}"]=b
        for col in ["A","B"]:
            ws[f"{col}{r}"].fill=C_IN; ws[f"{col}{r}"].border=BD; ws[f"{col}{r}"].alignment=A_WRAP; ws[f"{col}{r}"].font=F_N
    ws.column_dimensions["A"].width=28; ws.column_dimensions["B"].width=70
    ws.freeze_panes="A5"
    return wb


# ---------- framework parsing (turns an uploaded 2-column file into rows) ----------
def parse_framework(source):
    """
    source: a path or file-like .xlsx. Returns (framework, warnings).
    Rules:
      1. Find the header row (the row whose first two used cells read 'Competency' / 'Behaviour', any case).
         If none is found, assume data starts at row 1 in columns A/B.
      2. Read columns A (competency) and B (behaviour) below the header.
      3. Forward-fill: a blank competency inherits the last competency seen (grouped layouts).
      4. Skip rows with a blank behaviour, and skip fully blank rows.
      5. Trim whitespace. Reject if nothing valid is found.
    """
    wb = openpyxl.load_workbook(source, data_only=True)
    ws = wb["Competency Framework"] if "Competency Framework" in wb.sheetnames else wb.worksheets[0]

    def norm(v): return str(v).strip() if v is not None else ""

    # locate header row
    header_row = None
    for r in range(1, min(ws.max_row, 30) + 1):
        a = norm(ws.cell(r, 1).value).lower()
        b = norm(ws.cell(r, 2).value).lower()
        if a == "competency" and b == "behaviour":
            header_row = r
            break
    start = (header_row + 1) if header_row else 1

    framework, warnings = [], []
    last_comp = ""
    for r in range(start, ws.max_row + 1):
        comp = norm(ws.cell(r, 1).value)
        beh  = norm(ws.cell(r, 2).value)
        if comp:
            last_comp = comp
        if not beh:
            continue                      # skip blank-behaviour rows (incl. fully blank)
        if not last_comp:
            warnings.append(f"Row {r}: behaviour with no competency above it — skipped.")
            continue
        framework.append((last_comp, beh))

    if not framework:
        raise ValueError("No competency/behaviour rows found. Check the file has two columns "
                         "(Competency, Behaviour) with at least one filled row.")
    return framework, warnings



def parse_framework_text(text):
    """
    Parse a pasted framework. One behaviour per line, 'Competency, Behaviour'
    (comma or tab separated). A blank competency inherits the one above.
    A line with no separator is treated as a competency heading if it is the
    first thing seen, otherwise as a behaviour under the current competency.
    Returns (framework, warnings).
    """
    framework, warnings = [], []
    last_comp = ""
    for i, raw in enumerate(text.splitlines(), start=1):
        line = raw.strip()
        if not line:
            continue
        if line.lower() in ("competency\tbehaviour", "competency,behaviour", "competency behaviour"):
            continue  # header line
        if "\t" in line:
            parts = line.split("\t", 1)
        elif "," in line:
            parts = line.split(",", 1)
        else:
            parts = [line]
        if len(parts) == 2:
            comp, beh = parts[0].strip(), parts[1].strip()
            if comp:
                last_comp = comp
            if not beh:
                continue
            if not last_comp:
                warnings.append(f"Line {i}: behaviour with no competency yet — skipped.")
                continue
            framework.append((last_comp, beh))
        else:
            token = parts[0].strip()
            if not last_comp:
                last_comp = token          # first bare line = competency heading
            else:
                framework.append((last_comp, token))  # subsequent bare lines = behaviours
    if not framework:
        raise ValueError("Couldn't read any behaviours. Use one behaviour per line as "
                         "'Competency, Behaviour' (comma or tab between the two).")
    return framework, warnings


if __name__ == "__main__":
    # Running this file directly produces two example files so you can see it work.
    demo = [
        ("Drives Results", "Sets clear goals and holds the team accountable for outcomes."),
        ("Drives Results", "Removes obstacles and reprioritises to keep work moving toward targets."),
        ("Builds Collaboration", "Works across teams to solve shared problems and align on decisions."),
        ("Leads People", "Coaches team members with regular, specific feedback."),
    ]
    build_template().save("Framework_Template.xlsx")
    build_workbook(demo, ["OPQ", "MQ", "SJT", "Case Study", "Role Play", "BEI"]).save("Sample_Scoresheet.xlsx")
    print("Wrote Framework_Template.xlsx and Sample_Scoresheet.xlsx")
