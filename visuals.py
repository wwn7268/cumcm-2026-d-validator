"""Time-frequency matrices, independent of the optimization programs."""
from pathlib import Path
import numpy as np
from matplotlib.figure import Figure
from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib.colors import ListedColormap, BoundaryNorm
from matplotlib.patches import Patch
from matplotlib import rc_context

COLORS = ['#F3F5F7', '#2864A0', '#228572', '#DF9B32', '#8655B4', '#D63245']
LABELS = ['空闲', 'A类', 'B类', '原有C类', '新增C类', '冲突']
FONT = ['Microsoft YaHei', 'SimHei', 'DejaVu Sans']

def is_new(p):
    return bool(p.get('is_new')) or p.get('kind')=='new' or p.get('action')=='new' or str(p.get('id','')).upper().startswith(('NEW','CNEW'))

def matrices(plans, horizon):
    counts=np.zeros((horizon,100),dtype=np.int32)
    categories=np.zeros_like(counts,dtype=np.uint8)
    for p in plans:
        if p.get('canceled') or p.get('action')=='cancel':continue
        try:
            values=[p[k] for k in ('f0','f1','t0','t1','gap','count')]
            if any(isinstance(v,bool) or not isinstance(v,(int,float)) or not float(v).is_integer() for v in values):continue
            f0,f1,t0,t1,gap,n=map(int,values)
            duration=t1-t0
            if duration<=0 or f1<=f0 or gap<0 or not 0<=n<=10000:continue
            cat=4 if is_new(p) else {'A':1,'B':2,'C':3}.get(p.get('cls',str(p.get('id',''))[:1]),3)
            for r in range(n):
                a=t0+r*(duration+gap);b=a+duration
                lo,hi=max(0,f0),min(100,f1);left,right=max(0,a),min(horizon,b)
                if left<right and lo<hi:
                    counts[left:right,lo:hi]+=1
                    categories[left:right,lo:hi]=cat
        except (KeyError,TypeError,ValueError,OverflowError):
            continue
    categories[counts>1]=5
    return counts,categories

def make_figure(result):
    value=result.get('settings',{}).get('horizon',643)
    valid_horizon=isinstance(value,int) and not isinstance(value,bool) and 1<=value<=10000
    horizon=value if valid_horizon else 643
    counts,categories=matrices(result.get('plans',[]),horizon)
    with rc_context({'font.family':FONT,'axes.unicode_minus':False}):
        fig=Figure(figsize=(12.5,7.0),dpi=110,layout='constrained')
        FigureCanvasAgg(fig)
        ax1,ax2=fig.subplots(2,1,sharex=True,sharey=True)
        extent=(0,horizon,0,100)
        ax1.imshow(categories.T,origin='lower',extent=extent,aspect='auto',interpolation='nearest',
                   cmap=ListedColormap(COLORS),vmin=0,vmax=5)
        suffix='（可解析部分；判定未通过）' if not result.get('passed') else ''
        ax1.set_title('装备类别与冲突位置'+suffix,loc='left',fontsize=11,pad=10)
        ax1.legend(handles=[Patch(facecolor=c,label=l) for c,l in zip(COLORS,LABELS)],
                   loc='upper right',bbox_to_anchor=(1,1.20),ncol=6,frameon=False,fontsize=8)
        levels=max(2,int(counts.max()) if counts.size else 1)
        if levels<=12:
            from matplotlib import colormaps
            cmap=colormaps['YlOrRd'].resampled(levels+1)
            norm=BoundaryNorm(np.arange(-.5,levels+1.5),cmap.N)
            img=ax2.imshow(counts.T,origin='lower',extent=extent,aspect='auto',interpolation='nearest',cmap=cmap,norm=norm)
            fig.colorbar(img,ax=ax2,pad=.015,ticks=np.arange(levels+1),label='同时占用台数')
        else:
            img=ax2.imshow(counts.T,origin='lower',extent=extent,aspect='auto',interpolation='nearest',cmap='YlOrRd',vmin=0)
            fig.colorbar(img,ax=ax2,pad=.015,label='同时占用台数')
        ax2.set_title('占用计数矩阵（数值大于1表示冲突）',loc='left',fontsize=11)
        for ax in (ax1,ax2):
            ax.set_ylabel('频率 / Δf');ax.set_ylim(0,100);ax.set_xlim(0,horizon)
            ax.set_yticks([0,20,40,60,80,100])
        ax2.set_xlabel('时间 / Δt')
        fig.suptitle(f'问题{result.get("question", "")} · 时频资源矩阵',fontsize=13,fontweight='bold')
        note=f'半开区间；显示窗口 [0,{horizon}) × [0,100)。超界占用仅在图中裁剪，仍判为违规。'
        if not valid_horizon:note='时间窗输入不合法；此图仅为默认窗口的占位图，不代表完成验证。'
        fig.supxlabel(note,fontsize=8)
    return fig,counts,categories

def occupants(plans,t,f):
    found=[]
    for p in plans:
        if p.get('canceled') or p.get('action')=='cancel':continue
        try:
            if not p['f0']<=f<p['f1']:continue
            duration=p['t1']-p['t0'];step=duration+p['gap']
            if duration<=0 or step<=0 or t<p['t0']:continue
            r=int((t-p['t0'])//step)
            if r<p['count'] and t<p['t0']+r*step+duration:
                found.append(f'{p["id"]}（第{r+1}次）')
        except (KeyError,TypeError,ValueError):continue
    return found
