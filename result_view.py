"""将验证结果转换为界面所需的摘要，不改变判定或优化指标。"""
from __future__ import annotations


def _complete_originals(plans: list[dict]) -> bool:
    expected = {f"{cls}{n:03d}" for cls, count in (("A", 20), ("B", 40), ("C", 90))
                for n in range(1, count + 1)}
    required = {"id", "canceled", "df", "dt", "dg"}
    return (len(plans) == 150 and {p.get("id") for p in plans} == expected
            and all(required <= p.keys() for p in plans))


def build_result_view(result: dict) -> dict:
    """为四张摘要卡提供可信统计；未完整检查时用 None / — 表示未知。"""
    summary = result.get("summary", {})
    plans = result.get("plans", [])
    old = [p for p in plans if not p.get("is_new")]
    status = result.get("status", "invalid")
    complete = (status in {"feasible", "infeasible"}
                and summary.get("analysis_complete", False) and _complete_originals(old))
    if not complete:
        status = "invalid"

    view = dict(status=status, cancelled=None, adjusted=None, translation_distance=None,
                frequency_distance=None, time_distance=None, gap_distance=None,
                total_amplitude=None, added_count=None, notes=[])
    if complete:
        active = [p for p in old if not p["canceled"]]
        frequency = sum(abs(p["df"]) for p in active)
        time = sum(abs(p["dt"]) for p in active)
        gap = sum(abs(p["dg"]) for p in active)
        view.update(cancelled=sum(bool(p["canceled"]) for p in old),
                    adjusted=sum(any(p[k] for k in ("df", "dt", "dg")) for p in active),
                    translation_distance=frequency + time, frequency_distance=frequency,
                    time_distance=time, gap_distance=gap, total_amplitude=frequency + time + gap,
                    added_count=sum(bool(p.get("is_new")) and not p.get("canceled") for p in plans))

    if status == "feasible":
        view.update(title="验证通过", message="规则、边界及全部重复用频区间检查通过，未发现冲突。")
        check_value, check_hint = "全部", "满足当前验证条件"
    elif status == "infeasible":
        conflicts = len(result.get("conflicts", []))
        message = (f"发现 {conflicts} 次时频重叠，请查看下方冲突位置。" if conflicts
                   else "存在不符合题目规则的计划，请按下方提示修改。")
        view.update(title="验证失败", message=message)
        check_value, check_hint = "未通过", "具体原因见下方明细"
        view["notes"].append("以下为当前提交统计，不代表该方案有效。")
    else:
        view.update(title="验证未完成", message="附件未完整读入或检查未完成，暂不能给出完整统计。")
        check_value, check_hint = "未完成", "请先处理附件或检查错误"
        view["notes"].append("未完成检查的项目显示为“—”，不按 0 计。")

    scope = "第二问基础方案" if result.get("question") == 3 else "原有装备"
    distance_hint = (f"频移 {view['frequency_distance']} 步 · 时移 {view['time_distance']} 步"
                     if complete else "检查完成后统计")
    view["metrics"] = [
        dict(label="通过检查", value=check_value, hint=check_hint),
        dict(label="已撤销", value=str(view["cancelled"]) if complete else "—", hint=scope + " · 台"),
        dict(label="已调整", value=str(view["adjusted"]) if complete else "—", hint=scope + " · 台"),
        dict(label="总平移距离", value=str(view["translation_distance"]) if complete else "—", hint=distance_hint),
    ]
    if result.get("question") == 3:
        view["notes"].append("撤销、调整和平移统计均属于第二问基础方案；新增装备不计入这些指标。")
        if complete:
            view["notes"].append(f"本次附件新增 {view['added_count']} 台 C 类装备。")
    view["notes"].append("总平移距离为频率与首次时间的基本单位步数之和：Σ(|Δf|+|Δt|)。每台计一次，撤销及新增装备不计。")
    if result.get("question") == 4 and complete:
        gap_count = sum(bool(p["dg"]) for p in active)
        view["notes"].append(f"第四问已调整数包含 {gap_count} 台间隔调整；间隔变化量为 {view['gap_distance']} 步，单独计入总调整幅度 {view['total_amplitude']}，不计入总平移距离。")
    if status == "feasible":
        view["notes"].append("通过表示方案可行，不表示已达到最优。")
    return view
