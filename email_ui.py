"""Upload, edit, preview and download a financial-news email."""
from datetime import date
from hashlib import sha256

import streamlit as st

from email_formatter import GREETING, extract_upload, parse_report, plain_email, html_email, eml_email


def render_email_formatter():
    st.title('財金新聞信件自動排版')
    st.write('上傳 Word 或文字檔，依「開場問候 → 新聞標題 → 日期、來源與全文」排成信件。內容不摘要、不改寫。')
    st.caption('信件內文固定寬度 7.5 英吋（19.05 公分），預設標楷體 16 點。貼到 Outlook 時請選擇「保留來源格式」。')
    st.caption('支援 DOCX、TXT（每檔 10 MB 以內）。採文字排版，圖片與原檔的字型、顏色不會沿用。')
    uploaded = st.file_uploader('1. 上傳新聞檔案', type=['docx', 'txt'], key='email_upload')
    st.session_state.setdefault('email_source_text', '')
    st.session_state.setdefault('email_upload_signature', None)
    st.session_state.setdefault('email_upload_warnings', [])
    st.session_state.setdefault('email_prepared', None)
    if uploaded is not None:
        signature = sha256(uploaded.name.encode() + uploaded.getvalue()).hexdigest()
        if signature != st.session_state.email_upload_signature:
            try:
                text, warnings = extract_upload(uploaded.name, uploaded.getvalue())
            except ValueError as exc:
                st.error(str(exc))
                return
            st.session_state.email_source_text = text
            st.session_state.email_upload_signature = signature
            st.session_state.email_upload_warnings = warnings
            st.session_state.email_prepared = None
    else:
        st.session_state.email_upload_signature = None
    for warning in st.session_state.email_upload_warnings:
        st.warning(warning)
    text = st.text_area('2. 新聞內容（可直接貼上或修改）', key='email_source_text', height=280,
                        help='系統會依範本中的「本文共…字」辨識新聞標題。若沒有這些資訊，請在每篇標題前加上「## 」；Word 標題樣式也會自動辨識。')
    subject = st.text_input('3. 信件主旨', value=f'{date.today():%Y/%m/%d} 重要財金新聞匯集', key='email_subject')
    greeting = st.text_input('4. 開場問候', value=GREETING, key='email_greeting')
    with st.expander('排版選項'):
        font, size = '標楷體', 16
        st.caption('固定格式：全文標楷體 16 點，含標題、日期與來源；寬度 7.5 英吋。')
        clean = st.checkbox('移除「本文共…字」、播放時間及相鄰重複文字', value=False, key='email_clean',
                            help='預設保留全部文字。勾選後仍保留新聞日期、來源及內文；排版後會顯示移除的行數。')
    signature = sha256(repr(('kai16-width7.5-v2', text, subject, greeting, font, size, clean)).encode()).hexdigest()
    if st.button('產生信件預覽與下載', type='primary', key='email_generate'):
        try:
            report = parse_report(text, greeting, clean)
            st.session_state.email_prepared = dict(signature=signature, report=report,
                html=html_email(report, subject, font, size), plain=plain_email(report),
                eml=eml_email(report, subject, font, size))
        except ValueError as exc:
            st.session_state.email_prepared = None
            st.error(str(exc))
    result = st.session_state.email_prepared
    if result is None:
        return
    if result['signature'] != signature:
        st.info('內容或選項已變更，請重新產生信件，讓預覽與下載同步更新。')
        return
    report = result['report']
    st.success(f'已完成 {len(report.articles)} 篇新聞排版。' + (f' 已移除 {report.removed} 行。' if clean else ' 已保留原始文字。'))
    if report.inferred:
        st.info('未找到明確的分篇資訊，暫排為一篇。多篇新聞請在每篇標題前加上「## 」後重新產生。')
    st.download_button('下載信件草稿（EML）', result['eml'], '財金新聞信件.eml', mime='message/rfc822', on_click='ignore')
    st.download_button('下載保留格式版本（HTML）', result['html'].encode('utf-8'), '財金新聞信件_標楷體16點_7.5英吋.html', mime='text/html', on_click='ignore')
    st.download_button('下載純文字版本（TXT）', result['plain'].encode('utf-8-sig'), '財金新聞信件.txt', mime='text/plain', on_click='ignore')
    st.caption('EML 可由支援的郵件軟體開啟；是否直接進入編輯模式依軟體而定。也可開啟 HTML，選取信件內容後複製到 Outlook 或 Gmail。此功能只產生草稿，不會寄送。')
    st.subheader('信件預覽')
    st.text('主旨：' + subject)
    st.html(result['html'])
