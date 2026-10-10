#!/usr/bin/env python3
"""Unattended Chinese US-market weekly report; secrets via GitHub Actions only."""
import datetime as dt
import email
import imaplib
import json
import os
import re
import smtplib
import sys
import time
import traceback
import zipfile
from email.header import decode_header, make_header
from email.message import EmailMessage
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import yfinance as yf
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from openai import OpenAI

ROOT = Path("output")
ROOT.mkdir(exist_ok=True)
STATUS = ROOT / "status.json"
NY = ZoneInfo("America/New_York")
CORE = {"S&P 500": "^GSPC", "NASDAQ 100": "^NDX", "VGT": "VGT"}
WATCH = {
    "NVDA": "AI芯片", "AVGO": "定制AI芯片/网络", "AMD": "半导体",
    "MU": "存储", "SNDK": "NAND存储", "WDC": "存储", "STX": "存储",
    "DELL": "AI服务器", "SMCI": "AI服务器", "HPE": "服务器",
    "ANET": "AI网络", "MRVL": "半导体/网络", "MSFT": "云计算",
    "GOOGL": "云计算", "AMZN": "云计算", "ORCL": "云计算"
}

def record(state, **info):
    STATUS.write_text(json.dumps({"state": state, **info}, ensure_ascii=False, indent=2), encoding="utf-8")

def report_week():
    today = dt.datetime.now(NY).date()
    monday = today - dt.timedelta(days=today.weekday())
    if today.weekday() < 5:
        monday -= dt.timedelta(days=7)
    friday = monday + dt.timedelta(days=4)
    if dt.datetime.now(NY).date() == friday:
        raise RuntimeError("The Friday trading session has not necessarily closed.")
    return monday, friday

def daily_history(ticker, monday):
    start = monday - dt.timedelta(days=13)
    end = monday + dt.timedelta(days=7)
    hist = yf.Ticker(ticker).history(
        start=start.isoformat(), end=end.isoformat(), interval="1d",
        auto_adjust=True, actions=False
    )
    if hist.empty or "Close" not in hist:
        raise ValueError(f"No verified market series for {ticker}")
    hist = hist.copy()
    hist.index = pd.to_datetime(hist.index).date
    return hist

def summarize_prices(monday, friday):
    series = {}
    unavailable = {}
    for label, symbol in {**CORE, **{x:x for x in WATCH}}.items():
        try:
            h = daily_history(symbol, monday)
            before = h[h.index < monday]
            week = h[(h.index >= monday) & (h.index <= friday)]
            if before.empty or week.empty:
                raise ValueError("missing comparison sessions")
            if label in CORE and week.index.max() != friday:
                raise ValueError("Friday close missing; abort instead of stale weekly report")
            p0 = float(before.Close.iloc[-1])
            p1 = float(week.Close.iloc[-1])
            series[label] = {
                "ticker":symbol, "previous_close":round(p0,4),
                "week_close":round(p1,4), "week_change_pct":round((p1/p0-1)*100,2),
                "close_date":week.index[-1].isoformat(),
                "week_high":round(float(week.High.max()),4),
                "week_low":round(float(week.Low.min()),4),
                "week_volume":int(week.Volume.sum()),
            }
        except Exception as exc:
            unavailable[label] = str(exc)[:150]
    if any(x not in series for x in CORE):
        raise RuntimeError("Missing one or more core Friday closing prices: " + str(unavailable))
    ranked = sorted((k for k in WATCH if k in series),
                    key=lambda x:abs(series[x]["week_change_pct"]), reverse=True)
    return series, unavailable, ranked

def treasury_rates(monday, friday):
    # FRED publishes 2Y/10Y US Treasury constant maturity daily (observation dates).
    vals = {}
    try:
        for code in ("DGS2", "DGS10"):
            url = f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={code}"
            df = pd.read_csv(url)
            df["observation_date"] = pd.to_datetime(df["observation_date"]).dt.date
            df[code] = pd.to_numeric(df[code], errors="coerce")
            filtered = df[(df.observation_date >= monday - dt.timedelta(days=8)) &
                          (df.observation_date <= friday)].dropna(subset=[code])
            if not filtered.empty:
                r = filtered.iloc[-1]
                vals[code] = {"date":r.observation_date.isoformat(), "yield_pct":float(r[code])}
    except Exception as exc:
        vals["error"] = "Treasury series not obtained: " + str(exc)[:130]
    return vals

def research_sections(client, model, monday, friday, market, rates, ranking, missing):
    # Explicit data anchoring; a web-search model finds and cites current official releases.
    reference = json.dumps({
        "week_us_eastern": [monday.isoformat(), friday.isoformat()],
        "market_data":market, "treasury":rates,
        "candidate_momentum_rank":ranking[:12],
        "missing_data":missing,
        "price_data_source":"Yahoo Finance/yfinance adjusted daily closes; validate against official ETF/indices when possible"
    }, ensure_ascii=False)
    system = (
        "你是严谨的美国股票市场研究员。仅针对给定已结束的交易周写中文机构研究周报。"
        "先核查公开一手来源(SEC/公司IR/Fed/Treasury/BLS/BEA/ETF官网)，再补充权威新闻。"
        "价格以提供的数据为基础，明确指数点数/ETF股价的计量单位。"
        "绝不编造数据/新闻/财报/一致预期/期权资金。缺失必须写“未获得可靠数据”。"
        "每个重要事实附近标注可追踪网页 URL 和发布日期。"
        "事实、公司指引、市场一致预期、研究推断明确分开；不把相关性写成因果事实。"
        "每个章节都有机制推演和可证伪指标。"
        "使用 Markdown 标题 #/## 和表格，禁止给出具体买卖指令。"
    )
    part1 = (
        "写《美股深度周报》的第一至第六章，共约3500至5000汉字："
        "1 执行摘要5至8条；2 核心市场对象仅标普500、NASDAQ100、VGT，分析周涨跌、相对强弱、驱动、"
        "估值/利率、技术区间与两周催化；3 本周按日期行情传导链；4 宏观CPI/PPI/PCE/就业/ISM/FOMC"
        "和2Y/10Y美元原油，只写本周真实发布或影响本周的已发布数据，比较预期必须有来源；"
        "5 行业3至5组（AI芯片/服务器/存储/网络/云软件）；"
        "6 选动态4至8只股票深入分析，包含相对预期、订单/收入/毛利率/指引、价量、竞争、估值、风险。"
        "无法找到一致预期写明无可靠数据。每章附官方可追溯来源，数据保持准确。"
    )
    part2 = (
        "写《美股深度周报》的第七至十二章，共约2500至3500汉字："
        "7 当周财报和基本面，区分绝对增长和预期差；"
        "8 资金/市场宽度/期权/ETF资金流，有权威可靠来源才给定量；"
        "9 三大市场对象的日线周线技术风险，支撑阻力须基于给定价格区间；"
        "10 未来两周宏观/财报/行业日历，只有核实日期才能确定，说明预期差和传导路径；"
        "11 基准/偏多/偏空三情景，明确可检验触发因素；"
        "12 总结5个优先验证变量。"
        "最后给出独立“主要来源”清单，含可追溯 URL。"
        "充分研究，不杜撰、不用事后发生的信息冒充当时已知。"
    )
    outputs = []
    for title, instruction in [("第1—6章",part1),("第7—12章",part2)]:
        response = client.responses.create(
            model=model, tools=[{"type":"web_search"}],
            instructions=system,
            input=f"【输入行情数据】\n{reference}\n\n【写作任务：{title}】\n{instruction}",
            max_output_tokens=12500,
        )
        content = response.output_text or ""
        if len(content) < 2000:
            raise RuntimeError(f"Research output too short: {title}: {len(content)}")
        outputs.append(content)
    return "\n\n".join(outputs)

def write_docx(markdown, path, monday, friday):
    doc=Document()
    section=doc.sections[0]
    section.top_margin=Cm(1.8)
    section.bottom_margin=Cm(1.7)
    section.left_margin=Cm(2.0)
    section.right_margin=Cm(2.0)
    normal=doc.styles["Normal"]
    normal.font.name="Arial"
    normal.font.size=Pt(9.3)
    normal.paragraph_format.space_after=Pt(5)
    heading_font=RGBColor(28,54,87)
    for sty in ("Title","Heading 1","Heading 2","Heading 3"):
        doc.styles[sty].font.color.rgb=heading_font
    doc.add_heading("美股深度周报",0)
    doc.add_paragraph(f"美东交易周：{monday.isoformat()} 至 {friday.isoformat()}")
    doc.add_paragraph("研究报告｜事实与推断分别标示｜不构成证券买卖建议")
    for raw in markdown.splitlines():
        line=raw.strip()
        if not line:
            continue
        if line.startswith("# "):
            doc.add_heading(line[2:].strip(),1)
        elif line.startswith("## "):
            doc.add_heading(line[3:].strip(),2)
        elif line.startswith("### "):
            doc.add_heading(line[4:].strip(),3)
        elif re.match(r"^\|[\s:|\-]+\|$",line):
            continue
        else:
            # Markdown tables converted to compact text rows, to avoid losing values.
            line=re.sub(r"^\s*[-*]\s+", "• ",line)
            line=re.sub(r"\*\*(.*?)\*\*",r"\1",line)
            doc.add_paragraph(line)
    doc.save(path)
    if not path.exists() or path.stat().st_size<1500:
        raise RuntimeError("Word document absent/empty")
    with zipfile.ZipFile(path) as z:
        if z.testzip() is not None or "word/document.xml" not in z.namelist():
            raise RuntimeError("DOCX validation failed")

def compose_message(path, monday, friday, sender, recipient, preview):
    msg=EmailMessage()
    msg["From"]=sender
    msg["To"]=recipient
    msg["Subject"]=f"美股深度周报 {monday.isoformat()} 至 {friday.isoformat()}"
    msg["X-Report-Week"]=friday.isoformat()
    msg.set_content(
        f"本期美股深度周报统计区间为美东时间 {monday.isoformat()} 至 {friday.isoformat()}。\n"
        "Word 全文见附件，重点讨论标普500、纳指100、VGT，以及宏观利率、"
        "科技细分行业、动态个股、未来两周催化及条件情景。\n\n"
        "结论以附件执行摘要和来源为准。\n\n"
        f"附件文件：{path.name}\n"
    )
    msg.add_attachment(path.read_bytes(),maintype="application",
                       subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
                       filename=path.name)
    return msg

def recent_sent_has_attachment(user, password, subject, filename, since):
    with imaplib.IMAP4_SSL("imap.gmail.com",993,timeout=30) as box:
        box.login(user,password)
        # Gmail canonical mailbox often [Gmail]/Sent Mail; use XLIST special-use selection otherwise.
        status, folders=box.list()
        names=[]
        for row in folders or []:
            s=row.decode(errors="replace")
            if "\\Sent" in s or "Sent Mail" in s or "已发送" in s:
                names.append(s.split(' "/" ')[-1].strip('"'))
        names += ['"[Gmail]/Sent Mail"','"Sent"']
        selected=False
        for folder in names:
            try:
                status,_=box.select(folder,readonly=True)
                if status=="OK":
                    selected=True
                    break
            except Exception:
                continue
        if not selected:
            raise RuntimeError("Cannot locate Gmail Sent folder for verification")
        _,ids=box.search(None,'SINCE',since.strftime("%d-%b-%Y"))
        for num in (ids[0] or b"").split()[-120:][::-1]:
            status,parts=box.fetch(num,'(RFC822)')
            if status!="OK" or not parts or not isinstance(parts[0],tuple):
                continue
            item=email.message_from_bytes(parts[0][1])
            subj=str(make_header(decode_header(item.get("Subject",""))))
            if subj!=subject or item.get("To","").lower().find(
                os.environ["MAIL_TO"].lower())<0:
                continue
            for attachment in item.walk():
                name=attachment.get_filename()
                if name:
                    name=str(make_header(decode_header(name)))
                    if name==filename and len(attachment.get_payload(decode=True) or b"")>0:
                        return True
        return False

def main():
    monday,friday=report_week()
    path=ROOT / f"美股周报_{monday.isoformat()}_{friday.isoformat()}.docx"
    subject=f"美股深度周报 {monday.isoformat()} 至 {friday.isoformat()}"
    dry=os.environ.get("DRY_RUN","false").lower()=="true"
    api=os.environ.get("OPENAI_API_KEY","")
    user=os.environ.get("GMAIL_USER","")
    password=os.environ.get("GMAIL_APP_PASSWORD","").replace(" ","")
    recipient=os.environ.get("MAIL_TO","")
    if not api or not dry and not all([user,password,recipient]):
        raise RuntimeError("Missing required repository Secrets: OPENAI_API_KEY / GMAIL_USER / GMAIL_APP_PASSWORD / MAIL_TO")
    # Check before incurring LLM expense; if inaccessible, proceed but fail on post-send verification.
    if not dry:
        try:
            if recent_sent_has_attachment(user,password,subject,path.name,friday-dt.timedelta(days=1)):
                record("already_delivered",file=path.name)
                print("Matching DOCX attachment already sent; skipping.")
                return
        except Exception as exc:
            record("imap_precheck_unavailable",reason=str(exc)[:200])
    market,missing,ranking=summarize_prices(monday,friday)
    rates=treasury_rates(monday,friday)
    client=OpenAI(api_key=api,timeout=180)
    body=research_sections(client,os.environ.get("REPORT_MODEL","gpt-4.1"),
                           monday,friday,market,rates,ranking,missing)
    write_docx(body,path,monday,friday)
    record("report_ready",file=path.name,bytes=path.stat().st_size)
    if dry:
        record("dry_run_complete",file=path.name,bytes=path.stat().st_size)
        print("Dry run: report produced; no email sent.")
        return
    msg=compose_message(path,monday,friday,user,recipient,body[:500])
    with smtplib.SMTP_SSL("smtp.gmail.com",465,timeout=45) as smtp:
        smtp.login(user,password)
        smtp.send_message(msg)
    record("smtp_accepted_verification_pending",file=path.name,
           bytes=path.stat().st_size)
    for trial in range(3):
        time.sleep(12)
        try:
            if recent_sent_has_attachment(user,password,subject,path.name,friday-dt.timedelta(days=1)):
                record("verified_sent",file=path.name,bytes=path.stat().st_size)
                print("Sent-mail attachment verified; recipient inbox is not independently confirmed.")
                return
        except Exception as exc:
            record("smtp_accepted_imap_unavailable",reason=str(exc)[:200],file=path.name)
    # SMTP accepted: do not send twice, to avoid duplicate delivery.
    record("smtp_accepted_not_verified",file=path.name)
    raise RuntimeError("SMTP accepted mail but sent-folder attachment not verified. No automatic resend.")

if __name__=="__main__":
    try:
        main()
    except Exception as exc:
        print("REPORT_PIPELINE_FAILED:",type(exc).__name__,str(exc))
        if not STATUS.exists():
            record("failed",error=f"{type(exc).__name__}: {str(exc)}")
        sys.exit(1)
