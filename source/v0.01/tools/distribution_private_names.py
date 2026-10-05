"""Distribution-only private implementation names. No source/cache mismatch.

Only module-private constants and Nx-prefixed helpers are eligible; public
callbacks, procedure parameters, event handlers, attributes and all literals
remain unchanged. Both VBA source and compiled cache are built from this text.
"""
import re

TOKEN = re.compile(r'"(?:[^"]|"")*"|\x27[^\r\n]*|\bRem\b[^\r\n]*|\b[A-Za-z_][A-Za-z_0-9]*\b', re.I)
DECLARATION = re.compile(r'^Private\s+(?:Const\s+(NX_[A-Za-z_0-9]+)\b|(?:Function|Sub)\s+(Nx[A-Za-z_0-9]+)\b)', re.I|re.M)

def compact_private_names(source: str, global_literals: str = '') -> tuple[str, dict[str,str]]:
    candidates={next(value for value in match.groups() if value).lower() for match in DECLARATION.finditer(source)}
    strings='\n'.join(m.group() for m in TOKEN.finditer(source) if m.group().startswith('"'))+'\n'+global_literals
    public_lines='\n'.join(line for line in source.splitlines() if re.match(r'\s*(Public|Friend)\b',line,re.I))
    identifiers={m.group().lower() for m in TOKEN.finditer(source) if not m.group().startswith(('"',"'"))}
    mapping={}
    for name in sorted(candidates):
        if re.fullmatch(r'nxp[0-9]+x*',name):continue
        if re.search(r'\b'+re.escape(name)+r'\b',strings+'\n'+public_lines,re.I):continue
        if re.search(r'\b'+re.escape(name)+r'\s*:=',source,re.I):continue
        alias='nxp'+str(len(mapping)+1)
        while alias.lower() in identifiers:alias+='x'
        mapping[name]=alias
        identifiers.add(alias.lower())
    def replace(match):
        value=match.group()
        if value.startswith(('"',"'")) or re.match(r'^Rem\b',value,re.I):return value
        return mapping.get(value.lower(),value)
    return TOKEN.sub(replace,source),mapping
