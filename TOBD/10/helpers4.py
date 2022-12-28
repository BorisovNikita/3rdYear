import re
def f(tag: str) -> bool:
    return bool(re.match("\d+-\w+-or-less", str(tag)))
