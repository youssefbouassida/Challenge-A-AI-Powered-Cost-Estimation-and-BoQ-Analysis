import json

s = open("features.json", encoding="utf-8-sig").read()
print("total lines:", s.count("\n") + 1, "| total characters:", len(s))
try:
    json.loads(s)
    print("The file is valid JSON.")
except json.JSONDecodeError as e:
    print("Error:", e)
    print("--- text around the error (>>> marks the position) ---")
    print(repr(s[e.pos - 300:e.pos]) + "  >>>  " + repr(s[e.pos:e.pos + 150]))