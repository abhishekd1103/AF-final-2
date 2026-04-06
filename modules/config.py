"""config.py — Constants, level colors, column defs."""
import re

# Page sizes in DXA (1440 = 1 inch)
A4_W = 11906; A4_H = 16838  # portrait
A3_W = 16838; A3_H = 23811  # landscape for wide tables
MARGIN = 1134

# Threshold: >6 user columns (+ S.No.) = switch table section to landscape A3
WIDE_TABLE_THRESHOLD = 6

# Level colors from PPE table
LEVEL_COLORS = {
    "A":"E2EFDA","B":"E2EFDA","C":"D6E4F0","D":"FFF3CD",
    "E":"FFCDD2","F":"FFCDD2","G":"FFCDD2",
    "Level A":"E2EFDA","Level B":"E2EFDA","Level C":"D6E4F0","Level D":"FFF3CD",
    "Level E":"FFCDD2","Level F":"FFCDD2","Level G":"FFCDD2",
    "FCT Not Determined":"F2F2F2",
}
LEVEL_ORDER = {"A":1,"B":2,"C":3,"D":4,"E":5,"F":6,"G":7,
    "Level A":1,"Level B":2,"Level C":3,"Level D":4,
    "Level E":5,"Level F":6,"Level G":7}

AF_COLS = {
    "bus_id":       {"label":"Bus ID",                 "short":"Bus ID",    "req":True, "dtype":"text","round":False,"compare":False},
    "kv":           {"label":"Voltage (kV)",           "short":"kV",        "req":True, "dtype":"num", "round":False,"compare":False},
    "conductor_gap":{"label":"Conductor Gap (mm)",     "short":"Gap(mm)",   "req":False,"dtype":"num", "round":False,"compare":False},
    "working_dist": {"label":"Working Distance (cm)",  "short":"WD(cm)",    "req":False,"dtype":"num", "round":False,"compare":False},
    "rab":          {"label":"RAB (m)",                "short":"RAB(m)",    "req":False,"dtype":"num", "round":False,"compare":False},
    "glove_vrating":{"label":"Glove V-Rating (VAC)",   "short":"GloveV",   "req":False,"dtype":"num", "round":False,"compare":False},
    "glove_class":  {"label":"Glove Class",            "short":"Glove",     "req":True, "dtype":"text","round":False,"compare":False},
    "total_energy": {"label":"Incident Energy (cal/cm2)","short":"Energy",  "req":True, "dtype":"num", "round":True, "compare":True},
    "afb":          {"label":"Arc Flash Boundary (m)", "short":"AFB(m)",    "req":True, "dtype":"num", "round":True, "compare":True},
    "energy_level": {"label":"Energy Level",           "short":"Level",     "req":True, "dtype":"text","round":False,"compare":True},
    "final_fct":    {"label":"Clearing Time (sec)",    "short":"FCT(s)",    "req":False,"dtype":"num", "round":True, "compare":True},
    "source_pd":    {"label":"Source PD ID",           "short":"Source PD", "req":False,"dtype":"text","round":False,"compare":False},
    "total_ia":     {"label":"Arcing Current Ia (kA)", "short":"Ia(kA)",    "req":True, "dtype":"num", "round":True, "compare":True},
    "total_ibf":    {"label":"Bolted Fault Ibf (kA)",  "short":"Ibf(kA)",   "req":True, "dtype":"num", "round":True, "compare":True},
    "ppe_desc":     {"label":"PPE Description",        "short":"PPE Desc",  "req":False,"dtype":"text","round":False,"compare":False},
    "trip_time":    {"label":"Trip Time (sec)",        "short":"Trip(s)",   "req":False,"dtype":"num", "round":False,"compare":False},
    "open_time":    {"label":"Open Time (sec)",        "short":"Open(s)",   "req":False,"dtype":"num", "round":False,"compare":False},
    "total_pd_fct": {"label":"Total PD FCT (sec)",     "short":"PD FCT",    "req":False,"dtype":"num", "round":False,"compare":False},
}
ROUND_KEYS = [k for k,v in AF_COLS.items() if v["round"]]

AUTO_PATTERNS = {
    "bus_id":[r"^id$",r"bus.*id",r"bus.*name",r"^bus$",r"^name$"],
    "kv":[r"^kv$",r"voltage.*kv"],"conductor_gap":[r"conductor.*gap",r"gap.*ll"],
    "working_dist":[r"working.*dist",r"dist.*ll"],"rab":[r"^rab",r"restricted"],
    "glove_vrating":[r"glove.*v.*rat"],"glove_class":[r"glove.*class"],
    "total_energy":[r"total.*energy",r"incident.*energy",r"energy.*cal"],
    "afb":[r"^afb",r"arc.*flash.*bound"],"energy_level":[r"energy.*level",r"^level"],
    "final_fct":[r"final.*fct",r"clearing.*time"],"source_pd":[r"source.*pd"],
    "total_ia":[r"total.*ia",r"arcing.*current"],"total_ibf":[r"total.*ibf",r"bolted.*fault"],
    "ppe_desc":[r"ppe.*desc",r"^ppe$"],"trip_time":[r"trip.*time"],
    "open_time":[r"open.*time"],"total_pd_fct":[r"total.*pd.*fct"],
}
JUNK_PATTERNS = [
    r'^Project\s*:',r'^Location\s*:',r'^Contract\s*#',r'^Filename\s*:',
    r'^Engineer\s*:',r'^Date\s*:',r'^Serial\s*#',r'^ETAP',
    r'^Page\s+\d',r'^SKM',r'^EasyPower',r'^Copyright',
]
TABLE_MARKERS = {"PD_TABLE","COMPARISON_TABLE","SCENARIO_TABLE","SCENARIO_HEADING",
    "SCENARIO_DESC","SCENARIO_TABLE_TITLE","SCENARIO_END",
    "ANNEXURE_HEADING","ANNEXURE_CONTENT","ANNEXURE_END",
    "comparison_heading","comparison_description"}

def level_to_ppe(energy):
    try: e = float(energy)
    except: return "N/A","N/A"
    if e > 100: return "G","DO NOT WORK"
    if e > 40: return "F","DO NOT WORK"
    if e > 25: return "E","PPE Cat-4"
    if e > 8: return "D","PPE Cat-3"
    if e > 4: return "C","PPE Cat-2"
    if e > 2: return "B","PPE Cat-1"
    if e > 1.2: return "A","PPE Cat-1"
    return "Below A","Minimal"

def is_wide_table(n_user_cols):
    return n_user_cols > WIDE_TABLE_THRESHOLD

def sanitize_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '_', name).strip()

LANDSCAPE_COL_THRESHOLD = WIDE_TABLE_THRESHOLD
