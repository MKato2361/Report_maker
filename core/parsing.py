# report_maker/core/parsing.py
import re
from typing import Dict, Optional, Tuple
from datetime import datetime
from .settings import JST, WEEKDAYS_JA
from .textutil import normalize_text

LABEL_CANON = {
    "管理番号": "管理番号",
    "物件名": "物件名",
    "住所": "住所",
    "窓口会社": "窓口会社",
    "窓口": "窓口会社",
    "メーカー": "メーカー",
    "制御方式": "制御方式",
    "契約種別": "契約種別",
    "受信時刻": "受信時刻",
    "通報者": "通報者",
    "現着時刻": "現着時刻",
    "完了時刻": "完了時刻",
    "受信内容": "受信内容",
    "現着状況": "現着状況",
    "原因": "原因",
    "処置内容": "処置内容",
    "対応者": "対応者",
    "完了連絡先1": "完了連絡先1",
    "送信者": "送信者",
    "詳細はこちら": "受付URL",
    "現着・完了登録はこちら": "現着完了登録URL",
    "受付番号": "受付番号",
}
MULTILINE_KEYS = {"受信内容", "現着状況", "原因", "処置内容"}
LABEL_REGEX = re.compile(r"^\s*([^\s:：]+(?:・[^\s:：]+)?)\s*[:：]\s*(.*)$")

def _strip_url_tail(u: str) -> str:
    return re.sub(r"[)\]＞＞）」】>]+$", "", u.strip())

def try_parse_datetime(s: Optional[str]) -> Optional[datetime]:
    if not s:
        return None
    cand = s.strip().replace("年", "/").replace("月", "/").replace("日", "")
    cand = cand.replace("-", "/").replace("　", " ")
    for fmt in ("%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M", "%Y/%m/%d"):
        try:
            dt = datetime.strptime(cand, fmt)
            return dt.replace(tzinfo=JST)
        except Exception:
            pass
    return None

def split_dt_components(dt: Optional[datetime]) -> Tuple[Optional[int], Optional[int], Optional[int], Optional[str], Optional[int], Optional[int]]:
    if not dt:
        return None, None, None, None, None, None
    dt = dt.astimezone(JST)
    return dt.year, dt.month, dt.day, WEEKDAYS_JA[dt.weekday()], dt.hour, dt.minute

def minutes_between(a: Optional[str], b: Optional[str]) -> Optional[int]:
    s = try_parse_datetime(a); e = try_parse_datetime(b)
    if s and e:
        return int((e - s).total_seconds() // 60)
    return None

def first_date_yyyymmdd(*vals) -> str:
    for v in vals:
        dt = try_parse_datetime(v)
        if dt:
            return dt.strftime("%Y%m%d")
    from .settings import JST
    return datetime.now(JST).strftime("%Y%m%d")

def extract_fields(raw_text: str) -> Dict[str, Optional[str]]:
    t = normalize_text(raw_text)
    lines = t.split("\n")

    out_keys = {
        "管理番号","物件名","住所","窓口会社","メーカー","制御方式","契約種別",
        "受信時刻","通報者","現着時刻","完了時刻",
        "受信内容","現着状況","原因","処置内容",
        "対応者","送信者","受付番号","受付URL","現着完了登録URL",
        "作業時間_分","案件種別(件名)"
    }
    out: Dict[str, Optional[str]] = {k: None for k in out_keys}

    m_case = re.search(r"^件名:\s*【\s*([^】]+)\s*】", t, flags=re.MULTILINE)
    if m_case:
        out["案件種別(件名)"] = m_case.group(1).strip()
    m_mane = re.search(r"件名:.*?【[^】]+】\s*([A-Z0-9\-]+)", t, flags=re.IGNORECASE)
    subject_manageno = m_mane.group(1).strip() if m_mane else None

    current_multikey: Optional[str] = None
    buffer = []
    awaiting_url_for: Optional[str] = None

    def _flush_buffer():
        nonlocal buffer, current_multikey
        if current_multikey and buffer:
            val = "\n".join([ln for ln in buffer if ln.strip() != ""]).strip()
            out[current_multikey] = val or None
        buffer = []
        current_multikey = None

    i = 0
    while i < len(lines):
        line = lines[i]

        if awaiting_url_for and line.strip().startswith("http"):
            out[awaiting_url_for] = _strip_url_tail(line)
            awaiting_url_for = None
            i += 1
            continue

        m = LABEL_REGEX.match(line)
        if m:
            _flush_buffer()

            raw_label = m.group(1).strip()
            value_part = m.group(2).strip()
            canon = LABEL_CANON.get(raw_label)
            if canon is None:
                i += 1
                continue

            if canon in MULTILINE_KEYS:
                current_multikey = canon
                buffer = []
                if value_part:
                    buffer.append(value_part)
            elif canon in ("受付URL", "現着完了登録URL"):
                url = None
                if "http" in value_part:
                    murl = re.search(r"(https?://\S+)", value_part)
                    if murl:
                        url = _strip_url_tail(murl.group(1))
                if url:
                    out[canon] = url
                else:
                    awaiting_url_for = canon
            else:
                if canon == "管理番号" and not value_part and subject_manageno:
                    out[canon] = subject_manageno
                else:
                    out[canon] = value_part or out.get(canon)

            if "受付番号" in raw_label or "受付番号" in line:
                mnum = re.search(r"受付番号\s*[:：]\s*([0-9]+)", line)
                if mnum:
                    out["受付番号"] = mnum.group(1).strip()

            i += 1
            continue

        if current_multikey:
            buffer.append(line)
        else:
            if out.get("受付番号") is None:
                mnum = re.search(r"受付番号\s*[:：]\s*([0-9]+)", line)
                if mnum:
                    out["受付番号"] = mnum.group(1).strip()
        i += 1

    _flush_buffer()

    if not out.get("管理番号") and subject_manageno:
        out["管理番号"] = subject_manageno

    dur = minutes_between(out.get("現着時刻"), out.get("完了時刻"))
    out["作業時間_分"] = str(dur) if dur is not None and dur >= 0 else None
    return out
