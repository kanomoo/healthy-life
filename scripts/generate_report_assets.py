import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.sans-serif'] = ['Tahoma', 'Leelawadee UI']
plt.rcParams['axes.unicode_minus'] = False
import matplotlib.patches as patches
import numpy as np

os.makedirs('assets', exist_ok=True)

# 1. Generate diagram_data_protocol.png (Evidence Protocol Workflow)
def generate_protocol_diagram():
    fig, ax = plt.subplots(figsize=(10, 4.5), dpi=300)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 5)
    ax.axis('off')

    # Color palette
    c_blue = '#1E3A8A'
    c_teal = '#0D9488'
    c_amber = '#D97706'
    c_indigo = '#4338CA'
    c_bg = '#F8FAFC'
    c_border = '#CBD5E1'

    # Background card
    bg = patches.FancyBboxPatch((0.2, 0.2), 9.6, 4.6, boxstyle="round,pad=0.1,rounding_size=0.15",
                                facecolor=c_bg, edgecolor=c_border, linewidth=1.5)
    ax.add_patch(bg)

    # Title
    ax.text(5.0, 4.4, 'ระเบียบวิธีและกระบวนการเก็บรวบรวมหลักฐานเชิงประจักษ์ประจำวัน (Daily Evidence Protocol Workflow)',
            fontsize=12, fontweight='bold', ha='center', va='center', color='#0F172A', fontfamily='Tahoma')

    steps = [
        ("ขั้นตอนที่ 1\nก่อนเริ่มทำงาน", "ปรับสถานีงาน\n90-90-90\nตรวจสอบระดับสายตา\nบันทึกเวลาเริ่มใช้จอ", c_blue),
        ("ขั้นตอนที่ 2\nระหว่างการใช้จอ", "ตั้งเวลาสมาร์ตโฟน\nเตือนทุก 45 นาที\nลุกพักขยับตัว 2 นาที\nหมุนไหล่ช้าๆ 10 ครั้ง", c_teal),
        ("ขั้นตอนที่ 3\nหลังเลิกทำงาน", "บันทึก NRS ก่อนนวด\nนวดกดจุด 6 นาที\nกายบริหาร 6 ท่า\nบันทึก NRS หลังนวด", c_amber),
        ("ขั้นตอนที่ 4\nบันทึกข้อมูลคลาวด์", "บันทึกลงไฟล์ Excel\n6806021612037_\nHealth30D.xlsx\nซิงค์ผ่าน OneDrive", c_indigo),
    ]

    for idx, (title, desc, col) in enumerate(steps):
        x = 0.6 + idx * 2.3
        y = 1.0
        w = 1.9
        h = 2.8

        # Box
        box = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.08,rounding_size=0.12",
                                     facecolor='#FFFFFF', edgecolor=col, linewidth=2.0)
        ax.add_patch(box)

        # Header box
        hdr = patches.FancyBboxPatch((x, y + h - 0.7), w, 0.7, boxstyle="round,pad=0.08,rounding_size=0.12",
                                     facecolor=col, edgecolor=col, linewidth=1.0)
        ax.add_patch(hdr)

        ax.text(x + w/2, y + h - 0.35, title, fontsize=9.5, fontweight='bold', ha='center', va='center',
                color='#FFFFFF', fontfamily='Tahoma')
        ax.text(x + w/2, y + 1.0, desc, fontsize=8.5, ha='center', va='center',
                color='#1E293B', fontfamily='Tahoma', linespacing=1.4)

        # Arrow
        if idx < 3:
            ax.annotate('', xy=(x + w + 0.35, y + h/2), xytext=(x + w + 0.05, y + h/2),
                        arrowprops=dict(facecolor='#64748B', edgecolor='#64748B', width=2, headwidth=7, headlength=7))

    # Bottom milestone note
    ax.text(5.0, 0.5, '* การประเมินจุดสำคัญ (Milestones): บันทึกภาพถ่ายด้านข้างและวัดมุมข้อต่อด้วย ImageMeter ในวัน D1, D7, D14, D21, D30',
            fontsize=8.5, fontstyle='italic', ha='center', va='center', color='#475569', fontfamily='Tahoma')

    plt.tight_layout()
    out_path = 'assets/diagram_data_protocol.png'
    plt.savefig(out_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"Generated: {out_path}")

# 2. Generate excel_logbook_preview.png (Excel Grid Layout)
def generate_excel_preview():
    fig, ax = plt.subplots(figsize=(10, 4.2), dpi=300)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 5)
    ax.axis('off')

    # Window frame
    win = patches.FancyBboxPatch((0.2, 0.2), 9.6, 4.6, boxstyle="round,pad=0.05,rounding_size=0.08",
                                facecolor='#FFFFFF', edgecolor='#107C41', linewidth=2.0)
    ax.add_patch(win)

    # Excel Green Ribbon
    ribbon = patches.Rectangle((0.2, 4.1), 9.6, 0.7, facecolor='#107C41', edgecolor='none')
    ax.add_patch(ribbon)
    ax.text(0.5, 4.45, 'AutoSave ON', fontsize=8, color='#A7F3D0', fontweight='bold')
    ax.text(1.7, 4.45, '6806021612037_Health30D.xlsx - Microsoft Excel (OneDrive Sync)',
            fontsize=10.5, color='#FFFFFF', fontweight='bold')

    # Sheet tabs bar
    tabs = patches.Rectangle((0.2, 0.2), 9.6, 0.5, facecolor='#F3F4F6', edgecolor='#E5E7EB')
    ax.add_patch(tabs)
    tab1 = patches.Rectangle((0.5, 0.2), 2.2, 0.45, facecolor='#FFFFFF', edgecolor='#107C41', linewidth=1.5)
    ax.add_patch(tab1)
    ax.text(1.6, 0.42, 'Daily_Logbook (บันทึกรายวัน)', fontsize=8, color='#107C41', fontweight='bold', ha='center')
    ax.text(3.6, 0.42, 'Angle_Analysis (ภาพและมุม)', fontsize=8, color='#6B7280', ha='center')
    ax.text(5.6, 0.42, 'Summary_Metrics (สรุปผลและกราฟ)', fontsize=8, color='#6B7280', ha='center')

    # Table Header
    headers = ["Day", "Date", "Screen (Min)", "Break Target", "Actual Breaks", "NRS Pre", "NRS Post", "6 Exercises", "Status"]
    col_x = [0.4, 1.0, 2.0, 3.2, 4.4, 5.5, 6.4, 7.3, 8.6]
    widths = [0.6, 1.0, 1.2, 1.2, 1.1, 0.9, 0.9, 1.3, 1.0]

    # Header Row
    ax.add_patch(patches.Rectangle((0.3, 3.6), 9.4, 0.45, facecolor='#E2E8F0', edgecolor='#CBD5E1'))
    for h, x in zip(headers, col_x):
        ax.text(x, 3.82, h, fontsize=8, fontweight='bold', color='#1E293B', va='center')

    # Sample rows
    sample_data = [
        ("D1", "10/09/2569", "275", "6", "5", "6.0", "5.0", "Complete (6/6)", "Baseline Phase"),
        ("D2", "11/09/2569", "310", "6", "6", "6.0", "5.0", "Complete (6/6)", "Active Break 100%"),
        ("D7", "16/09/2569", "420", "9", "8", "5.0", "4.0", "Complete (6/6)", "Milestone 1 Pass"),
        ("D14", "23/09/2569", "480", "10", "9", "4.0", "3.0", "Complete (6/6)", "Midterm Pass (NRS -33%)"),
        ("...", "...", "...", "...", "...", "...", "...", "...", "..."),
        ("AVG", "Days 1-14", "350.6 min", "7.3 times", "6.5 times", "4.64", "3.64", "85.7% Compliant", "Target Achieved "),
    ]

    for r_idx, row in enumerate(sample_data):
        y_pos = 3.15 - r_idx * 0.42
        bg_col = '#F8FAFC' if r_idx % 2 == 0 else '#FFFFFF'
        if r_idx == len(sample_data) - 1:
            bg_col = '#DCFCE7'
        ax.add_patch(patches.Rectangle((0.3, y_pos - 0.12), 9.4, 0.4, facecolor=bg_col, edgecolor='#E2E8F0', linewidth=0.5))
        for val, x in zip(row, col_x):
            weight = 'bold' if r_idx == len(sample_data) - 1 else 'normal'
            color = '#15803D' if r_idx == len(sample_data) - 1 else '#334155'
            ax.text(x, y_pos + 0.08, val, fontsize=7.5, fontweight=weight, color=color, va='center')

    plt.tight_layout()
    out_path = 'assets/excel_logbook_preview.png'
    plt.savefig(out_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"Generated: {out_path}")

# 3. Generate imagemeter_angle_measurement.png
def generate_imagemeter_preview():
    fig, ax = plt.subplots(figsize=(8, 4.5), dpi=300)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 6)
    ax.axis('off')

    # Card background
    ax.add_patch(patches.FancyBboxPatch((0.2, 0.2), 9.6, 5.6, boxstyle="round,pad=0.1,rounding_size=0.12",
                                        facecolor='#0F172A', edgecolor='#334155', linewidth=1.5))

    # Top title bar
    ax.text(0.6, 5.3, 'ImageMeter - Mobile Measurement Tool (3-Point Ergonomic Joint Angle Analysis)',
            fontsize=10.5, color='#38BDF8', fontweight='bold')

    # Draw stick-figure sitting posture with key angles
    # Hip at (3.5, 2.0), Shoulder at (3.4, 4.0), Elbow at (4.5, 3.2), Wrist at (5.5, 3.2)
    # Knee at (5.2, 2.0), Ankle at (5.2, 0.8), Head at (3.4, 4.6), Screen at (6.5, 3.8)

    # Desk & Chair
    ax.plot([5.3, 7.5], [3.1, 3.1], color='#64748B', linewidth=4) # desk
    ax.plot([2.5, 4.2], [1.9, 1.9], color='#64748B', linewidth=4) # chair seat
    ax.plot([2.8, 2.8], [1.9, 4.2], color='#64748B', linewidth=4) # chair back

    # Torso (Shoulder to Hip)
    ax.plot([3.4, 3.5], [4.0, 2.0], color='#F1F5F9', linewidth=3)
    # Thigh (Hip to Knee)
    ax.plot([3.5, 5.2], [2.0, 2.0], color='#F1F5F9', linewidth=3)
    # Lower leg (Knee to Ankle)
    ax.plot([5.2, 5.2], [2.0, 0.8], color='#F1F5F9', linewidth=3)
    # Upper arm (Shoulder to Elbow)
    ax.plot([3.4, 4.3], [4.0, 3.0], color='#F1F5F9', linewidth=3)
    # Forearm (Elbow to Wrist)
    ax.plot([4.3, 5.4], [3.0, 3.2], color='#F1F5F9', linewidth=3)
    # Head
    circle = patches.Circle((3.4, 4.6), 0.35, facecolor='#F1F5F9', edgecolor='none')
    ax.add_patch(circle)

    # Eye level line
    ax.plot([3.6, 6.5], [4.6, 4.6], color='#F59E0B', linestyle='--', linewidth=1.5)
    ax.text(6.6, 4.6, 'Eye Level (Pass)', color='#F59E0B', fontsize=8, fontweight='bold', va='center')

    # Points and angle annotations
    # 1. Elbow angle at (4.3, 3.0)
    ax.plot(4.3, 3.0, 'ro', markersize=6)
    ax.text(4.1, 2.5, 'Elbow: 98°\n(Target 90-100° )', color='#4ADE80', fontsize=8, fontweight='bold')

    # 2. Hip angle at (3.5, 2.0)
    ax.plot(3.5, 2.0, 'ro', markersize=6)
    ax.text(2.1, 2.3, 'Hip: 108°\n(Target 90-100°)', color='#FACC15', fontsize=8, fontweight='bold')

    # 3. Knee angle at (5.2, 2.0)
    ax.plot(5.2, 2.0, 'ro', markersize=6)
    ax.text(5.4, 1.6, 'Knee: 92°\n(Target 85-100° )', color='#4ADE80', fontsize=8, fontweight='bold')

    # Right side legend / measurement spec
    ax.add_patch(patches.FancyBboxPatch((6.8, 0.8), 2.7, 3.4, boxstyle="round,pad=0.08,rounding_size=0.08",
                                        facecolor='#1E293B', edgecolor='#475569', linewidth=1.0))
    ax.text(7.0, 3.9, '3-Point Angle Specs:', color='#FFFFFF', fontsize=8.5, fontweight='bold')
    ax.text(7.0, 3.4, '1. Elbow (ไหล่-ศอก-ข้อมือ)\n   Baseline: 138° -> D14: 98° ', color='#94A3B8', fontsize=7.5)
    ax.text(7.0, 2.7, '2. Hip (ไหล่-สะโพก-เข่า)\n   Baseline: 125° -> D14: 108°', color='#94A3B8', fontsize=7.5)
    ax.text(7.0, 2.0, '3. Knee (สะโพก-เข่า-ข้อเท้า)\n   Baseline: 65° -> D14: 92° ', color='#94A3B8', fontsize=7.5)
    ax.text(7.0, 1.2, '4. Overall Score: 3/4 (75%)\n   (Passed Target ≥75% )', color='#4ADE80', fontsize=8, fontweight='bold')

    plt.tight_layout()
    out_path = 'assets/imagemeter_angle_measurement.png'
    plt.savefig(out_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"Generated: {out_path}")

if __name__ == '__main__':
    generate_protocol_diagram()
    generate_excel_preview()
    generate_imagemeter_preview()
