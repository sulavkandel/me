#!/usr/bin/env python3
"""Generate a YIC-2026 tailored CV for Sulav Kandel as a clean one/two-page PDF."""

from fpdf import FPDF

NAVY = (23, 42, 70)
ACCENT = (40, 90, 150)
GRAY = (90, 90, 90)
LINE = (200, 200, 200)


class CV(FPDF):
    def header(self):
        pass

    def footer(self):
        pass


def section_title(pdf, text):
    pdf.ln(2)
    pdf.set_font("Helvetica", "B", 11.5)
    pdf.set_text_color(*NAVY)
    pdf.cell(0, 6, text.upper(), new_x="LMARGIN", new_y="NEXT")
    y = pdf.get_y()
    pdf.set_draw_color(*ACCENT)
    pdf.set_line_width(0.5)
    pdf.line(pdf.l_margin, y, pdf.w - pdf.r_margin, y)
    pdf.ln(1.5)


def role(pdf, title, place, date, bullets):
    pdf.set_font("Helvetica", "B", 10.5)
    pdf.set_text_color(20, 20, 20)
    pdf.cell(0, 5.2, title, new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "I", 9.5)
    pdf.set_text_color(*GRAY)
    line2 = place
    if date:
        line2 = f"{place}  |  {date}"
    pdf.cell(0, 5, line2, new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 9.8)
    pdf.set_text_color(35, 35, 35)
    for b in bullets:
        pdf.set_x(pdf.l_margin + 1)
        pdf.cell(4, 4.8, chr(149))
        pdf.multi_cell(0, 4.8, b, new_x="LMARGIN", new_y="NEXT")
    pdf.ln(1.5)


pdf = CV(format="A4")
pdf.set_auto_page_break(auto=True, margin=12)
pdf.set_margins(15, 12, 15)
pdf.add_page()

# Name
pdf.set_font("Helvetica", "B", 22)
pdf.set_text_color(*NAVY)
pdf.cell(0, 9, "SULAV KANDEL", new_x="LMARGIN", new_y="NEXT")
pdf.set_font("Helvetica", "", 10.5)
pdf.set_text_color(*ACCENT)
pdf.cell(0, 5.5, "Engineering Student  |  Tech-for-Good Builder  |  Community Volunteer", new_x="LMARGIN", new_y="NEXT")
pdf.set_font("Helvetica", "", 9.3)
pdf.set_text_color(*GRAY)
pdf.cell(0, 5.5, "Kathmandu, Nepal  |  +977 9845384841  |  connectwithsulav@gmail.com", new_x="LMARGIN", new_y="NEXT")
pdf.ln(1)

# Profile
section_title(pdf, "Profile")
pdf.set_font("Helvetica", "", 9.8)
pdf.set_text_color(35, 35, 35)
profile = ("Final year Electronics, Communication and Information Engineering student from Nepal who builds "
           "technology with a social purpose. Experience spans live telecom network optimization, blockchain "
           "systems for fair elections, and applied machine learning. Equally active off-campus as an AI skills "
           "instructor and a volunteer with women's health and empowerment groups. Driven by a simple belief that "
           "technology should serve people first, and keen to connect with young changemakers solving global problems.")
pdf.multi_cell(0, 4.9, profile, new_x="LMARGIN", new_y="NEXT")

# Education
section_title(pdf, "Education")
role(pdf, "B.E. in Electronics, Communication and Information Engineering", "Tribhuvan University, IOE Paschimanchal Campus, Pokhara", "2021 - Apr 2026",
     ["Coursework across telecommunications, signal processing, programming, and applied machine learning."])
role(pdf, "Higher Secondary Education (10+2), Science", "Kathmandu Model College, Bagbazar, Kathmandu", "2018 - 2020", [])

# Leadership & Community
section_title(pdf, "Leadership and Community Impact")
role(pdf, "Instructor, AI Content Creation", "Nepal Skill Development Center, Kathmandu", "Present",
     ["Teach hands-on classes on AI tools for video, image, and script creation to learners from varied backgrounds.",
      "Focus on making new technology accessible to people who would otherwise be left behind by it."])
role(pdf, "Executive Member", "Nurture Nari (NPO), Women's Health Awareness", "Jan 2026 - Present",
     ["Help plan and run community awareness activities focused on women's health in Nepal."])
role(pdf, "Member", "Subharambha Nepal, Women Empowerment Community", "2023 - Present",
     ["Support community programs that promote education and empowerment for women and girls."])
role(pdf, "Student Ambassador", "Fonepay Nepal", "2025",
     ["Led campus outreach campaigns and activations to grow awareness of digital payment products among students."])

# Experience
section_title(pdf, "Professional Experience")
role(pdf, "Intern, GSM / LTE RF Operations", "Nepal Telecom (NTC), GSM Department, Pokhara", "Jan 2026 - Mar 2026",
     ["Studied 2G/3G/4G-LTE network architecture and RF planning across live cellular sites.",
      "Performed RF drive testing, KPI measurement, and log analysis to evaluate signal quality and coverage.",
      "Carried out handover analysis and applied network optimization concepts on operational cells."])
role(pdf, "Content Creator (Tech, Finance, History)", "NewsBreak (Remote)", "Jan 2022 - Feb 2023",
     ["Researched trending topics and produced written and short-form content for online audiences."])

# Projects
section_title(pdf, "Selected Projects")
role(pdf, "SolVote - Blockchain-Based E-Voting System", "Solana, React.js, Node.js, SQL, Raspberry Pi", "2024",
     ["Co-developed a decentralized voting framework aimed at fair, tamper-proof, one-person-one-vote elections.",
      "Combined RFID and biometric authentication with NFT-based voting tokens for verified voters.",
      "Integrated Raspberry Pi hardware with a React.js frontend and Node.js backend for real-time results."])
role(pdf, "GNN-DQN Handover Optimization and Load Balancing", "Python, PyTorch, GNN, Reinforcement Learning", "2025-26",
     ["Built a Deep Q-Network agent with a Graph Neural Network to optimize handover and load balancing in cellular networks.",
      "Designed reward functions to cut ping-pong handovers, improve throughput, and balance cell usage."])

# Skills
section_title(pdf, "Skills and Achievements")
pdf.set_font("Helvetica", "B", 9.8)
pdf.set_text_color(20, 20, 20)
pdf.cell(28, 4.8, "Technical:")
pdf.set_font("Helvetica", "", 9.8)
pdf.set_text_color(35, 35, 35)
pdf.multi_cell(0, 4.8, "Python, C/C++, JavaScript, SQL; React.js, Node.js, PyTorch, Pandas; ML (GNN, DQN, RL), Telecom (RF planning, drive testing), Blockchain (Solana), IoT (Raspberry Pi).", new_x="LMARGIN", new_y="NEXT")
pdf.set_font("Helvetica", "B", 9.8)
pdf.set_text_color(20, 20, 20)
pdf.cell(28, 4.8, "Highlights:")
pdf.set_font("Helvetica", "", 9.8)
pdf.set_text_color(35, 35, 35)
pdf.multi_cell(0, 4.8, "Finalist, Sun Securities Stock Pitching Competition (2026); Participant, Pokhara Metropolitan City Hackathon (2026).", new_x="LMARGIN", new_y="NEXT")
pdf.set_font("Helvetica", "B", 9.8)
pdf.set_text_color(20, 20, 20)
pdf.cell(28, 4.8, "Languages:")
pdf.set_font("Helvetica", "", 9.8)
pdf.set_text_color(35, 35, 35)
pdf.multi_cell(0, 4.8, "Nepali (native), English (professional working proficiency).", new_x="LMARGIN", new_y="NEXT")

pdf.output("Sulav_Kandel_CV_YIC2026.pdf")
print("CV written: Sulav_Kandel_CV_YIC2026.pdf")
