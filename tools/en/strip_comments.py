import re, shutil, sys, zipfile
src, dst = sys.argv[1], sys.argv[2]
zin = zipfile.ZipFile(src)
drop = [n for n in zin.namelist()
        if n.startswith("ppt/comments/") or n == "ppt/authors.xml"]
dropset = set(drop)
zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    n = it.filename
    if n in dropset:
        continue
    data = zin.read(n)
    if n.endswith(".rels"):
        t = data.decode("utf-8")
        t = re.sub(r'<Relationship[^>]*Target="[^"]*(?:comments/[^"]*|authors\.xml)"[^>]*/>', "", t)
        data = t.encode("utf-8")
    elif re.match(r"ppt/slides/slide\d+\.xml$", n):
        t = data.decode("utf-8")
        t = re.sub(r"<\w+:commentRel[^>]*/>", "", t)
        t = re.sub(r"<\w+:commentRel[^>]*>.*?</\w+:commentRel>", "", t, flags=re.S)
        data = t.encode("utf-8")
    elif n == "[Content_Types].xml":
        t = data.decode("utf-8")
        t = re.sub(r'<Override[^>]*PartName="/ppt/(?:comments/[^"]*|authors\.xml)"[^>]*/>', "", t)
        data = t.encode("utf-8")
    zout.writestr(it, data)
zout.close()
print("removed", len(drop), "comment parts")
