import json, itertools, random
words = ["BALLADE","LYRIK","STROPHE","VERS","REIM","PAARREIM","KREUZREIM","UMARMENDER","KEHRREIM","EPIK","ERZÄHLER","DRAMATIK","DIALOG"]
def try_build(order, seed):
    random.seed(seed)
    grid = {}; placed = []
    def can(w, r, c, d):
        dr, dc = (0,1) if d=="H" else (1,0)
        cross = 0
        # cell before/after must be empty
        if grid.get((r-dr, c-dc)) or grid.get((r+dr*len(w), c+dc*len(w))): return -1
        for i,ch in enumerate(w):
            rr, cc = r+dr*i, c+dc*i
            g = grid.get((rr,cc))
            if g:
                if g != ch: return -1
                cross += 1
            else:
                # neighbours perpendicular must be empty
                if dr==0:
                    if grid.get((rr-1,cc)) or grid.get((rr+1,cc)): return -1
                else:
                    if grid.get((rr,cc-1)) or grid.get((rr,cc+1)): return -1
        return cross
    def put(w,r,c,d):
        dr, dc = (0,1) if d=="H" else (1,0)
        for i,ch in enumerate(w): grid[(r+dr*i, c+dc*i)] = ch
        placed.append((w,r,c,d))
    put(order[0],0,0,"H")
    for w in order[1:]:
        opts=[]
        for (pw,pr,pc,pd) in placed:
            for i,ch in enumerate(pw):
                for j,ch2 in enumerate(w):
                    if ch!=ch2: continue
                    if pd=="H": r,c,d = pr-j, pc+i, "V"
                    else: r,c,d = pr+i, pc-j, "H"
                    k = can(w,r,c,d)
                    if k>=1: opts.append((k,r,c,d))
        opts=[o for o in opts if (o[1],o[2]) not in [(pr,pc) for _,pr,pc,_ in placed]]
        if not opts: return None
        random.shuffle(opts); opts.sort(key=lambda x:-x[0])
        put(w,*opts[0][1:])
    rs=[r for r,c in grid]; cs=[c for r,c in grid]
    return placed, max(rs)-min(rs)+1, max(cs)-min(cs)+1, min(rs), min(cs)
best=None
for seed in range(4000):
    order=[words[0]]+random.Random(seed).sample(words[1:],len(words)-1) if seed%2 else random.Random(seed).sample(words,len(words))
    res=try_build(order,seed)
    if not res: continue
    placed,h,w,r0,c0=res
    score=h*w + abs(h-w)*3
    if best is None or score<best[0]: best=(score,placed,h,w,r0,c0)
score,placed,h,w,r0,c0=best
placed=[(wd,r-r0,c-c0,d) for wd,r,c,d in placed]
print(h,w); print(placed)
json.dump({"h":h,"w":w,"placed":placed},open("kw.json","w"))
g=[["." for _ in range(w)] for _ in range(h)]
for wd,r,c,d in placed:
    for i,ch in enumerate(wd):
        if d=="H": g[r][c+i]=ch
        else: g[r+i][c]=ch
print("\n".join(" ".join(x) for x in g))
