"""Outlook COM helpers — build HTML mail with CID-embedded PNGs, save to Drafts."""
import os
from pathlib import Path
from tempfile import NamedTemporaryFile

import win32com.client as win32

from agent.logging_setup import get_logger

PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
OL_MAIL_ITEM = 0
OL_FORMAT_HTML = 2


def create_draft(
    to: str,
    subject: str,
    html_body: str,
    cid_pngs: dict | None = None,
) -> str:
    """Create an Outlook mail item, attach PNGs as CID inline images, save to Drafts.

    cid_pngs maps CID strings (e.g. 'mtd_chart') to raw PNG bytes. The HTML body
    must reference them via <img src='cid:mtd_chart'>.

    Returns the draft's EntryID (useful for logging or later retrieval).
    Does NOT call .Send() — explicit reviewer step.
    """
    log = get_logger()
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(OL_MAIL_ITEM)
    mail.To = to
    mail.Subject = subject
    mail.BodyFormat = OL_FORMAT_HTML

    temp_paths: list[str] = []
    try:
        if cid_pngs:
            for cid, png_bytes in cid_pngs.items():
                tmp = NamedTemporaryFile(delete=False, suffix=".png")
                tmp.write(png_bytes)
                tmp.close()
                temp_paths.append(tmp.name)
                att = mail.Attachments.Add(tmp.name)
                att.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, cid)
                log.debug("Attached CID image: %s (%d bytes)", cid, len(png_bytes))

        mail.HTMLBody = html_body
        mail.Save()
        entry_id = mail.EntryID
        log.info("Outlook draft saved (EntryID=%s)", entry_id)
        return entry_id
    finally:
        for p in temp_paths:
            try:
                os.unlink(p)
            except OSError:
                pass
