"""
export_chat_transcript.py
Parses the session transcript and generates a clean, readable Markdown archive:
CHAT_TRANSCRIPT.md
"""

import json
from pathlib import Path
from datetime import datetime

def export_transcript():
    base_dir = Path(__file__).resolve().parent
    transcript_file = Path("/Users/deep/.gemini/antigravity/brain/4d192bca-25e3-4b70-bc3f-ed1d0ff4fe13/.system_generated/logs/transcript_full.jsonl")
    if not transcript_file.exists():
        transcript_file = Path("/Users/deep/.gemini/antigravity/brain/4d192bca-25e3-4b70-bc3f-ed1d0ff4fe13/.system_generated/logs/transcript.jsonl")
        
    out_md = base_dir / "CHAT_TRANSCRIPT.md"
    
    dialogues = []
    
    with open(transcript_file, "r", encoding="utf-8") as f:
        for line in f:
            if not line.strip():
                continue
            entry = json.loads(line)
            source = entry.get("source")
            msg_type = entry.get("type")
            content = entry.get("content", "")
            created_at = entry.get("created_at", "")
            
            # Format timestamp
            time_str = ""
            if created_at:
                try:
                    dt = datetime.fromisoformat(created_at.replace("Z", "+00:00"))
                    time_str = dt.strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    time_str = created_at
            
            # Only include genuine user inputs and model responses
            if source == "USER_EXPLICIT" and msg_type == "USER_INPUT":
                # Clean up XML tags if any
                clean_content = content
                if "<USER_REQUEST>" in clean_content:
                    parts = clean_content.split("<USER_REQUEST>")
                    if len(parts) > 1:
                        clean_content = parts[1].split("</USER_REQUEST>")[0].strip()
                if clean_content:
                    dialogues.append({
                        "speaker": "Deep Koshiya (User)",
                        "time": time_str,
                        "text": clean_content
                    })
            elif source == "MODEL" and content.strip():
                # Avoid logging simple internal updates if redundant
                if content.strip().startswith("[Message]") or content.strip().startswith("<SYSTEM_MESSAGE>"):
                    continue
                dialogues.append({
                    "speaker": "AI Senior Analytics Engineer & Data Scientist",
                    "time": time_str,
                    "text": content.strip()
                })

    # Build Markdown
    lines = [
        "# Complete Project Discussion & Chat Transcript",
        "",
        "**Project:** E-Commerce Customer Lifecycle & RFM Segmentation Dashboard  ",
        "**Candidate / Author:** Deep Koshiya  ",
        "**Academic Program:** B.Sc. in Artificial Intelligence & Data Science (Semester V)  ",
        "**Institute:** Institute of Advanced Research (IAR University)  ",
        "**Milestone:** First Progress Work Evaluation (24/09/2026)  ",
        "",
        "---",
        "",
        "## Session Overview",
        "This document contains the complete chronological dialogue, architectural consultations, technical blueprints, code specifications, report formatting iterations, and milestone alignments recorded during the development of this project.",
        "",
        "---",
        ""
    ]

    for idx, d in enumerate(dialogues):
        speaker = d["speaker"]
        time_lbl = f" *({d['time']})*" if d["time"] else ""
        if "Deep Koshiya" in speaker:
            lines.append(f"### 👤 {speaker}{time_lbl}")
            lines.append("")
            lines.append("> " + d["text"].replace("\n", "\n> "))
            lines.append("")
        else:
            lines.append(f"### 🤖 {speaker}{time_lbl}")
            lines.append("")
            lines.append(d["text"])
            lines.append("")
        lines.append("---")
        lines.append("")

    with open(out_md, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
        
    print(f"[SUCCESS] Exported complete chat transcript with {len(dialogues)} interactions to {out_md}")

if __name__ == "__main__":
    export_transcript()
