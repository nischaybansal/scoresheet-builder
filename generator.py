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
SJT_SCALES = ["Managerial Judgement","Managing Objectives","People Management","Reputation Management",
"Big Picture","Delegative","One-to-One","Team","Personal Recognition","Company Protocol"]

# ---- tool registry: canonical order = online first, then offline ----
# kind: 'online' (paste-stens), 'bars' (rubric anchors), 'bei' (interview questions)
# score_col = the column on the tool sheet holding the behaviour's 1-5 rating that DI reads
VERIFY_DIMS = ["Inductive","Deductive","Numerical"]

TOOLS = {
    "OPQ":              {"kind":"online","dims":OPQ_DIMS,   "family":"online"},
    "MQ":               {"kind":"online","dims":MQ_DIMS,    "family":"online"},
    "SJT":              {"kind":"online","dims":SJT_SCALES, "family":"online"},
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
ONLINE_AVG_COL = "L"   # competency-average column on online sheets (feeds Detailed Integration)

HR = 6          # header row on every sheet
FB = HR + 1     # first behaviour row (behaviour i -> row FB+i on every tool sheet)

# ---- styling ----
ARIAL="Montserrat"   # all cells use Montserrat 9 (must be installed for Excel to render it)
C_HDR=PatternFill("solid",fgColor="1F3864"); C_SUB=PatternFill("solid",fgColor="D9E1F2")
C_IN=PatternFill("solid",fgColor="FFF2CC");  C_LOCK=PatternFill("solid",fgColor="E2EFDA")
C_LEG=PatternFill("solid",fgColor="FCE4D6"); C_PASTE=PatternFill("solid",fgColor="DDEBF7")
C_DIMBG=PatternFill("solid",fgColor="F2F2F2")
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
    ws=wb.create_sheet(name)
    _title(ws,f"{name}  —  online psychometric",
        "SCORING KEY (constant): Positive trait -> Score = roundup(Sten/2): 1-2=1 · 3-4=2 · 5-6=3 · 7-8=4 · 9-10=5.  "
        "Negative trait -> inverse: 1-2=5 · 3-4=4 · 5-6=3 · 7-8=2 · 9-10=1.","L")
    ws["A4"]=("Assessor pastes stens in the blue block on the right. Consultant maps up to "
              f"{TRAIT_SLOTS} traits per competency, writes the Rating 1-5 rubric descriptions, and sets "
              "Direction (Positive by default). Sten, Score and the Competency Average fill automatically.")
    ws["A4"].font=F_IT; ws.merge_cells("A4:L4"); ws["A4"].alignment=A_WV
    heads=["Competency","Behaviour",f"Mapped {name} Trait","Direction",
           "Rating 1","Rating 2","Rating 3","Rating 4","Rating 5","Sten","Score (1-5)","Competency Avg"]
    cols=["A","B","C","D","E","F","G","H","I","J","K","L"]
    for c,h in zip(cols,heads): ws[f"{c}{HR}"]=h
    _hdr(ws,HR,cols)
    # paste-stens block on the right
    ws[f"N{HR}"]=f"{name} Dimension"; ws[f"O{HR}"]="Paste Sten"; _hdr(ws,HR,["N","O"])
    for i,dim in enumerate(dims):
        r=FB+i
        ws[f"N{r}"]=dim; ws[f"N{r}"].font=F_N; ws[f"N{r}"].border=BD; ws[f"N{r}"].fill=C_DIMBG
        ws[f"O{r}"]=""; ws[f"O{r}"].fill=C_PASTE; ws[f"O{r}"].border=BD; ws[f"O{r}"].alignment=A_CTR
    dim_last=FB+len(dims)-1
    geo=_block_geometry(framework)
    for comp,behs,gs,ge in geo:
        for k in range(ge-gs+1):
            r=gs+k
            if k<len(behs):
                ws[f"A{r}"]=comp; ws[f"B{r}"]=behs[k]
            if k<TRAIT_SLOTS:
                ws[f"C{r}"]=""; ws[f"C{r}"].fill=C_IN
                ws[f"D{r}"]="Positive"; ws[f"D{r}"].fill=C_IN
                for rc in ["E","F","G","H","I"]:           # Rating 1-5 rubric cells (consultant fills)
                    ws[f"{rc}{r}"]=""; ws[f"{rc}{r}"].fill=C_SUB
                ws[f"J{r}"]=f'=IFERROR(INDEX($O${FB}:$O${dim_last},MATCH(C{r},$N${FB}:$N${dim_last},0)),"")'
                ws[f"J{r}"].fill=C_LOCK
                ws[f"K{r}"]=(f'=IF(J{r}="","",IF(D{r}="Negative",6-ROUNDUP(J{r}/2,0),ROUNDUP(J{r}/2,0)))')
                ws[f"K{r}"].fill=C_LOCK
            _cellsetup(ws,r,cols)
            for cc in ["D","J","K"]: ws[f"{cc}{r}"].alignment=A_CTR
        g=ws[f"{ONLINE_AVG_COL}{gs}"]
        g.value=f'=IFERROR(ROUND(AVERAGE(K{gs}:K{ge}),0),"")'
        g.fill=C_LOCK; g.font=F_B; g.alignment=A_CTR
        if ge>gs: ws.merge_cells(f"{ONLINE_AVG_COL}{gs}:{ONLINE_AVG_COL}{ge}")
        for rr in range(gs,ge+1): ws[f"{ONLINE_AVG_COL}{rr}"].border=BD
    last=geo[-1][3]
    dv=DataValidation(type="list",formula1=f"=$N${FB}:$N${dim_last}",allow_blank=True)
    ws.add_data_validation(dv); dv.add(f"C{FB}:C{last}")
    dv2=DataValidation(type="list",formula1='"Positive,Negative"',allow_blank=True)
    ws.add_data_validation(dv2); dv2.add(f"D{FB}:D{last}")
    for c,w in zip(["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O"],
                   [20,38,22,10,17,17,17,17,17,6,9,12,3,22,10]):
        ws.column_dimensions[c].width=w
    ws.freeze_panes=f"A{FB}"

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

# ---------- Detailed Integration ----------
def _build_di(wb,selected,framework):
    di=wb.create_sheet("Detailed Integration")
    di["A1"]="DETAILED INTEGRATION"; di["A1"].font=F_TITLE
    for lab,r in [("Participant Name",2),("Assessor Name",3),("Date of Scoring",4)]:
        di[f"A{r}"]=lab; di[f"A{r}"].font=F_B; di[f"B{r}"]=""; di[f"B{r}"].fill=C_IN; di[f"B{r}"].border=BD
    fixed=["Competency","Behaviour","Behaviour Rating","Competency Rating"]
    tool_cols=[t for t in CANONICAL if t in selected]   # online-first ordering
    heads=fixed+tool_cols
    cols=[openpyxl.utils.get_column_letter(i+1) for i in range(len(heads))]
    for c,h in zip(cols,heads): di[f"{c}{HR}"]=h
    _hdr(di,HR,cols)
    # ONLINE/OFFLINE tags above the tool columns
    for c,t in zip(cols[4:],tool_cols):
        di[f"{c}{HR-1}"]=TOOLS[t]["family"].upper(); di[f"{c}{HR-1}"].font=F_TAG; di[f"{c}{HR-1}"].alignment=A_CTR
    start=FB; idx=0; di_row=start; groups=[]
    beh_col=cols[2]; comp_rating_col=cols[3]
    has_online=any(TOOLS[t]["family"]=="online" for t in tool_cols)
    online_geo=_block_geometry(framework) if has_online else []
    gi=0
    for comp,items in groupby(framework,key=lambda x:x[0]):
        items=list(items); gs=di_row
        for (c,beh) in items:
            trow=FB+idx
            di[f"{cols[0]}{di_row}"]=comp; di[f"{cols[1]}{di_row}"]=beh
            di[f"{beh_col}{di_row}"]=""; di[f"{beh_col}{di_row}"].fill=C_IN
            for k,t in enumerate(tool_cols):
                if TOOLS[t]["family"]=="online":
                    continue  # online tools land once per competency (below)
                dcol=cols[4+k]; sc=TOOLS[t]["score_col"]
                di[f"{dcol}{di_row}"]=f'=IF({_ref(t)}!{sc}{trow}="","",{_ref(t)}!{sc}{trow})'
                di[f"{dcol}{di_row}"].fill=C_LOCK
            _cellsetup(di,di_row,cols)
            for cc in cols[2:]: di[f"{cc}{di_row}"].alignment=A_CTR
            idx+=1; di_row+=1
        ge=di_row-1
        # online tools: one competency-level average per block, merged down the group
        blk_start=online_geo[gi][2] if online_geo else None
        for k,t in enumerate(tool_cols):
            if TOOLS[t]["family"]!="online":
                continue
            dcol=cols[4+k]
            cell=di[f"{dcol}{gs}"]
            cell.value=f'=IF({_ref(t)}!{ONLINE_AVG_COL}{blk_start}="","",{_ref(t)}!{ONLINE_AVG_COL}{blk_start})'
            cell.fill=C_LOCK; cell.font=F_B; cell.alignment=A_CTR
            if ge>gs: di.merge_cells(f"{dcol}{gs}:{dcol}{ge}")
            for rr in range(gs,ge+1): di[f"{dcol}{rr}"].border=BD
        gi+=1
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

def build_workbook(framework, selected_tools):
    """framework: list of (competency, behaviour). selected_tools: subset of CANONICAL."""
    selected=[t for t in CANONICAL if t in selected_tools]
    framework=_cluster_by_competency(framework)   # ensures per-competency averaging
    wb=openpyxl.Workbook(); wb.remove(wb.active)
    for name in selected:
        spec=TOOLS[name]
        if spec["kind"]=="online": _build_online(wb,name,spec["dims"],framework)
        elif spec["kind"]=="bars": _build_bars(wb,name,framework)
        elif spec["kind"]=="bei":  _build_bei(wb,name,framework)
    _build_di(wb,selected,framework)
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
