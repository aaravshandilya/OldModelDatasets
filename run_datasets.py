#!/usr/bin/env python3
"""Run the supplied Java generator for all five dataset families. Requires Java 17+."""
from pathlib import Path
import csv, subprocess, time, os, zipfile
ROOT=Path(__file__).resolve().parent
CONFIGS=[('berlin',10,1001),('vienna',20,2001),('london',50,3001),('paris',100,4001),('milan',200,5001)]
def main():
    cp=os.pathsep.join(str(p) for p in sorted((ROOT/'deps').glob('*.jar')))
    subprocess.run(['java','com.sun.tools.javac.Main','-cp',cp,str(ROOT/'source/Main.java')],check=True)
    rows=[]
    for city,n,seed in CONFIGS:
        folder=ROOT/f'{city}_{n}_hubs'; folder.mkdir(exist_ok=True)
        start=time.perf_counter()
        result=subprocess.run(['java','-Xmx1g','-cp',str(ROOT/'source')+os.pathsep+cp,'Main',str(folder)],input=f'{n}\n{n}\n3\n1\n10\n{seed}\n',text=True,capture_output=True)
        for isn in range(seed, seed+10):
            workbook=folder/f'{isn}.xlsx'
            if not zipfile.is_zipfile(workbook):
                retry=subprocess.run(['java','-Xmx1g','-cp',str(ROOT/'source')+os.pathsep+cp,'Main',str(folder)],input=f'{n}\n{n}\n3\n1\n1\n{isn}\n',text=True,capture_output=True,check=True)
                result.stdout += f'\nRegenerated incomplete workbook {isn}.xlsx\n' + retry.stdout + retry.stderr
                if not zipfile.is_zipfile(workbook): raise RuntimeError(f'Invalid workbook: {workbook}')
        elapsed=time.perf_counter()-start
        (folder/'generation.log').write_text(result.stdout+result.stderr)
        result.check_returncode()
        if 'EXCEPTION CAUGHT' in result.stdout: raise RuntimeError(result.stdout)
        for rep in range(1,11):
            isn=seed+rep-1
            rows.append(dict(city=city,hubs=n,replicate=rep,i=n,j=n,m=3,k=1,seed=isn,workbook=f'{folder.name}/{isn}.xlsx',csv_folder=f'{folder.name}/{isn}',Eik_rows=n,Njk_rows=n,Bmjk_rows=3*n,Cijk_rows=n*n,batch_seconds=round(elapsed,4)))
        print(f'{city}: generated 10 datasets, {n} hubs, {elapsed:.2f} s',flush=True)
    with (ROOT/'dataset_index.csv').open('w',newline='') as f:
        writer=csv.DictWriter(f,fieldnames=list(rows[0]));writer.writeheader();writer.writerows(rows)
if __name__=='__main__': main()
