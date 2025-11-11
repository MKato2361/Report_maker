def extract_fields(raw_text: str) -> Dict[str, Optional[str]]:
    t = normalize_text(raw_text)

    # 件名由来の補助抽出
    subject_case = _search_one(r"件名:\s*【\s*([^】]+)\s*】", t, flags=re.IGNORECASE)
    subject_manageno = _search_one(r"件名:.*?【[^】]+】\s*([A-Z0-9\-]+)", t, flags=re.IGNORECASE)

    # 1行想定（行末まで）
    single_line = {
        "管理番号": r"(?im)^\s*管理番号\s*:\s*([A-Za-z0-9\-]+)\s*$",
        "物件名": r"(?im)^\s*物件名\s*:\s*(.+)$",
        "住所": r"(?im)^\s*住所\s*:\s*(.+)$",
        "窓口会社": r"(?im)^\s*窓口\s*:\s*(.+)$",
        "メーカー": r"(?im)^\s*メーカー\s*:\s*(.+)$",
        "制御方式": r"(?im)^\s*制御方式\s*:\s*(.+)$",
        "契約種別": r"(?im)^\s*契約種別\s*:\s*(.+)$",
        "受信時刻": r"(?im)^\s*受信時刻\s*:\s*([0-9/\-:\s]+)$",
        "現着時刻": r"(?im)^\s*現着時刻\s*:\s*([0-9/\-:\s]+)$",
        "完了時刻": r"(?im)^\s*完了時刻\s*:\s*([0-9/\-:\s]+)$",
        "通報者": r"(?im)^\s*通報者\s*:\s*(.+)$",
        "対応者": r"(?im)^\s*対応者\s*:\s*(.+)$",
        # 「完了連絡先1」「完了連絡先」どちらでもOK
        "完了連絡先1": r"(?im)^\s*完了連絡先(?:1)?\s*:\s*(.+)$",
        "送信者": r"(?im)^\s*送信者\s*:\s*(.+)$",
        "受付番号": r"(?im)^\s*受付番号\s*:\s*([0-9]+)\s*$",
        "受付URL": r"(?im)^\s*詳細はこちら\s*:\s*.*?(https?://\S+)\s*$",
        "現着完了登録URL": r"(?im)^\s*現着・完了登録はこちら\s*:\s*(https?://\S+)\s*$",
    }

    # 複数行想定（本文ブロックのみをスパン抽出）
    multiline_labels = {
        "受信内容": r"受信内容\s*:",
        "現着状況": r"現着状況\s*:",
        "原因": r"原因\s*:",
        "処置内容": r"処置内容\s*:",
    }

    out: Dict[str, Optional[str]] = {k: None for k in set(single_line.keys()) | set(multiline_labels.keys())}
    out.update({
        "案件種別(件名)": subject_case,
        "受付URL": None,
        "現着完了登録URL": None,
    })

    # 行単位でまず拾う
    for k, pat in single_line.items():
        out[k] = _search_one(pat, t, flags=re.IGNORECASE)

    # 件名からの管理番号補完
    if not out.get("管理番号") and subject_manageno:
        out["管理番号"] = subject_manageno

    # 本文ブロックだけをスパン抽出（次の本文ラベルまで）
    for k in multiline_labels:
        span = _search_span_between(multiline_labels, k, t)
        if span:
            out[k] = span

    # ---- 仕上げのガード：行頭アンカーでクリーンに再取得（巻き込み防止）----
    # 対応者
    m = re.search(r"(?im)^\s*対応者\s*:\s*(.+)$", t)
    if m:
        out["対応者"] = m.group(1).strip()

    # 通報者
    m = re.search(r"(?im)^\s*通報者\s*:\s*(.+)$", t)
    if m:
        out["通報者"] = m.group(1).strip()

    # 完了連絡先1（= 完了連絡先 も可）
    m = re.search(r"(?im)^\s*完了連絡先(?:1)?\s*:\s*(.+)$", t)
    if m:
        out["完了連絡先1"] = m.group(1).strip()

    # 作業時間（分）
    dur = minutes_between(out.get("現着時刻"), out.get("完了時刻"))
    out["作業時間_分"] = str(dur) if dur is not None and dur >= 0 else None

    return out
