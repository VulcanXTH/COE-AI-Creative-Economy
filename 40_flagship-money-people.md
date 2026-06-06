# 40 — รายละเอียดราย Flagship: ประเมินเงิน + คน

> **ขอบเขต:** ประเมินทรัพยากร "เงิน (งบประมาณ)" และ "คน (กำลังคน/FTE)" ราย 5 โครงการเรือธง
> **ฐานข้อมูล:** ผลวิเคราะห์ bottom-up ราย Agent (ไฟล์ 36) + แผนงบ 500 (ไฟล์ 37) · หน่วย: ล้านบาท (ลบ.) · 5 มิ.ย. 2569
> **อ่านอย่างไร:** แต่ละโครงการแสดง **(ก) งบเต็มรูปแบบ (bottom-up)** = ทรัพยากรที่ต้องใช้จริงถ้าทำเต็ม · **(ข) ในแผน 500** = ฉบับลดขนาดที่จะเริ่มจริง · ส่วน "คน" แสดงทีมเต็มรูปแบบ (แผน 500 ลดทีมตามสัดส่วน)

---

## ภาพรวมเงิน + คน ทั้ง 5 โครงการ

| โครงการ | งบเต็ม (bottom-up) | ในแผน 500 | ทีมหลัก (FTE peak) | ทักษะวิกฤต |
|---|---:|---:|---:|---|
| 1. Foundation Model | 517 *(Fine-tune)* | 120 | ~16 core + 25 annotate | นักวิจัย AI, Thai NLP, MLOps |
| 2. Bangkok AI Summit | ~329 | 80 | ~6 core + PCO จ้างเหมา | event/MICE, business matching |
| 3. Sandbox Platform | 424 | 150 | 6 → 9 | สถาปนิกระบบ, MLOps, mentor |
| 4. Creator AI Toolkit | 457 | 100 | ~15–20 (รวม content/mentor) | curriculum, edtech, mentor |
| 5. Export Accelerator | 296 | 30 | 6 → 9 | trade/market, accelerator |
| **รวม** | **~2,023** | **500** | **~50–70 core** (+ pool) | |

---

## 🏛 โครงการที่ 1 — Thai Creative AI Foundation Model

### 💰 เงิน
**งบเต็ม (Fine-tune): 517 ลบ. / 5 ปี** · ในแผน 500: **120 ลบ. (Dataset-first + POC, เลื่อน train ใหญ่)**

| องค์ประกอบ | งบเต็ม (Fine-tune) | หมายเหตุ |
|---|---:|---|
| ทีมวิจัย/วิศวกร AI | 165 | ก้อนใหญ่สุด (คน) |
| Compute/GPU (LANTA in-kind ลด 30%) | 78 | ไม่ใช่ตัวขับหลัก |
| ข้อมูลไทย: จัดเตรียม+annotate | 78 | ทีม annotate 25 คน |
| Data/Content licensing | 45 | **จุดเสี่ยงค่าใช้จ่าย** |
| โครงสร้าง/Storage + Eval + Serving + Safety | 96 | |
| Overhead 12% | 55 | |

> ⚠️ จุดชี้ขาดเงิน: **ข้อมูล+ลิขสิทธิ์ (123 ลบ.) > compute (78 ลบ.)** — ต้นทุนจริงคือ "ข้อมูลสร้างสรรค์ไทย" ไม่ใช่ GPU

### 👥 คน (ทีมเต็มรูปแบบ Fine-tune)
| บทบาท | FTE | เงินเดือน/เดือน |
|---|---:|---|
| Lead/Principal Researcher | 2 | 180,000 |
| วิศวกร AI/ML | 8 | 100,000 |
| Data/MLOps Engineer | 4 | 100,000 |
| PM/ประสานงาน | 2 | 90,000 |
| **ทีม annotate ข้อมูล (ช่วง Data phase)** | +25 | 35,000 |

- **ทักษะวิกฤต:** นักวิจัย AI อาวุโส (หายากในไทย), Thai NLP, MLOps, การจัดการลิขสิทธิ์ข้อมูล
- **กำลังคนตามเฟส:** P0 8–10 → P1 (data) 20–25 → P2 (train) 18–22 → P3 (eval) 10–12 → P4 (release) 12–15 → P5 10
- **บทบาทหน่วยงาน:** BDI นำ data+ธรรมาภิบาล · สวทช./NECTEC นำวิจัย+LANTA · AICAT ดึงเครือข่ายศิลปิน/เจ้าของเนื้อหา + human eval
- **ความเสี่ยงคน:** ขาดบุคลากร AI อาวุโส → ร่วม ม./ดึงคนไทยตปท./advisor นานาชาติ + ถ่ายทอดจาก Typhoon/SEA-LION

---

## 🎪 โครงการที่ 2 — Bangkok AI Creative Summit

### 💰 เงิน
**งบเต็ม (เกาะ event → standalone ปี 3): ~329 ลบ.** · ในแผน 500: **80 ลบ. (เกาะ event ตลอด)**

| รูปแบบ | ต่อครั้ง (net) | องค์ประกอบหลัก |
|---|---:|---|
| เกาะ event เดิม | ~27 ลบ. | เพิ่มเฉพาะ 1 เวที + matching + speaker (incremental) |
| Standalone สากล | ~82–95 ลบ. | venue 4.8 + production/AV 18 + keynote 8.4 + การตลาด 9 + PCO 11 + อื่นๆ |

> ตัวขับเงิน: production/AV + keynote ต่างชาติ (USD 60K/คน tier-1) + venue · เป้า cost-recovery จาก sponsor 35–45%

### 👥 คน (โมเดล "ทีมเล็ก + จ้างเหมา PCO")
| บทบาท | รูปแบบ |
|---|---|
| Content/Program Lead | ทีมภายใน (เนื้อหา/agenda) |
| Business Matching Lead | DITP (เครือข่าย buyer/ทูตพาณิชย์) |
| Operations/Logistics | จ้างเหมา PCO + สสปน. (MICE) |
| Marketing/PR | ทีมภายใน + agency |

- **ทักษะวิกฤต:** บริหารงาน MICE ระดับสากล, business matchmaking, การดึง speaker/buyer ต่างชาติ
- **กำลังคน:** ทีมแกนภายใน ~6 คน + **outsource งานผลิตทั้งหมดให้ PCO** (ไม่ต้องจ้างทีมใหญ่ประจำ)
- **บทบาทหน่วยงาน:** DITP เจ้าภาพ (buyer/matching) · CEA (เนื้อหาสร้างสรรค์) · AICAT (เนื้อหาเทคนิค) · สสปน./TCEB (MICE ops + เงินสนับสนุน)
- **ความเสี่ยงคน:** พึ่ง PCO → ต้องมีสัญญา+KPI ชัด, ทีมภายในคุมทิศทาง/ดีลไม่ใช่ปล่อย PCO ทั้งหมด

---

## 🧪 โครงการที่ 3 — AI Creative Sandbox Platform

### 💰 เงิน
**งบเต็ม: 424 ลบ. / 5 ปี** (860 โปรเจกต์) · ในแผน 500: **150 ลบ.** (~300–400 โปรเจกต์)

| องค์ประกอบ | งบเต็ม | สัดส่วน |
|---|---:|---:|
| Compute/Cloud credits | 150.5 | **36% (ตัวขับหลัก)** |
| ทีมบริหารโครงการ (FTE) | 61.5 | 15% |
| พัฒนาแพลตฟอร์ม (CapEx) | 45.0 | 11% |
| ดูแล+ข้อมูล+license+mentor+คัด/ประเมิน+outreach | ~129 | 30% |
| Contingency 10% | 38.6 | |

> คุมเงิน: **cap credit/โปรเจกต์ + ใช้ spot GPU** (compute เป็น 36%)

### 👥 คน (6 → 9 FTE)
| บทบาท | ปี1 | ปลายแผน |
|---|---:|---:|
| PM/หัวหน้าโครงการ | 1 | 1 |
| สถาปนิกระบบ/แพลตฟอร์ม | 1 | 1 |
| MLOps/วิศวกร provisioning | 1 | 2 |
| Mentor/ผู้ประเมินเทคนิค (pool) | 2 | 3 |
| คัดเลือก/ประเมินผล/baseline + admin | 1 | 2 |
| **รวม** | **6** | **9** |

- **ทักษะวิกฤต:** สถาปัตยกรรมแพลตฟอร์ม, MLOps/FinOps (คุม compute), mentor เทคนิค, ประเมิน commercialization
- **บทบาทหน่วยงาน:** NIA เจ้าภาพ (คัดเลือก+commercialization) · BDI (data governance) · AICAT (mentor เทคนิค) · depa (voucher engine)
- **ความเสี่ยงคน:** mentor คุณภาพไม่พอ → สร้าง mentor pool จากเครือข่าย AICAT/อุตสาหกรรม

---

## 🎓 โครงการที่ 4 — Thai Creator Economy AI Toolkit

### 💰 เงิน
**งบเต็ม: 457 ลบ. / 5 ปี** (600k ผู้เรียน) · ในแผน 500: **100 ลบ.** (~200k ผู้เรียน)

| องค์ประกอบ | งบเต็ม | สัดส่วน |
|---|---:|---:|
| License เครื่องมือ AI (pool seats) | 189.0 | **45% (ตัวขับหลัก)** |
| หลักสูตร/เนื้อหา/MOOC | 64.3 | 14% |
| รับรองสมรรถนะ (สคช.) | 53.5 | 12% |
| วิทยากร/ครูพี่เลี้ยง + แพลตฟอร์ม + การตลาด + PMO | ~150 | 29% |

> คุมเงิน: **pool seats หมุนเวียน + เจรจาราคา EDU/ภาครัฐ** (ห้ามให้ 1 seat/คน 600k คน = พุ่งหลักพันล้าน)

### 👥 คน (ทีมหลัก + pool ขยายตามผู้ใช้)
| บทบาท | FTE |
|---|---:|
| PMO/บริหารโครงการ | 4 |
| ทีมเทคนิค/แพลตฟอร์ม | 3 |
| Curriculum Lead + ทีมผลิตเนื้อหา | 2 + 6–8 |
| Mentor/ครูพี่เลี้ยง (pool) | 15–20 |
| Marketing/รับสมัคร | 4 |

- **ทักษะวิกฤต:** ออกแบบหลักสูตร (instructional design), edtech/LMS, mentor, เจรจา license vendor
- **บทบาทหน่วยงาน:** depa (งบ/จัดซื้อ license + LMS) · CEA (เนื้อหาสร้างสรรค์) · AICAT (องค์ความรู้ AI) · สคช. (มาตรฐาน+รับรอง) · มหาวิทยาลัย (ผลิต MOOC + วิทยากร)
- **ความเสี่ยงคน:** ทีมผลิตเนื้อหาไม่ทันการ refresh AI → ใช้ rapid authoring + AI narration + มหาวิทยาลัยร่วม

---

## 🌐 โครงการที่ 5 — AI Creative Export Accelerator

### 💰 เงิน
**งบเต็ม: 296 ลบ. / 5 ปี** (170 บริษัท) · ในแผน 500: **30 ลบ.** (นำร่องเล็ก fold เข้า Summit)

| องค์ประกอบ | งบเต็ม |
|---|---:|
| Pavilion งานต่างประเทศ | 81.0 |
| Accelerator/mentoring | 76.5 |
| Trade mission/คณะผู้แทนการค้า | 49.5 |
| PMO + market intelligence + แพลตฟอร์มจับคู่ + contingency | ~89 |

> คุมเงิน: **cap ขนาด pavilion ≤250 ตร.ม. + co-funding สสว.** (อุดหนุน 50–80% บริษัทร่วมจ่าย) · บทเรียน Thai Pavilion Osaka ~868 ลบ.

### 👥 คน (PMO 6 → 9 FTE)
| บทบาท | ปี1 | ปลายแผน |
|---|---:|---:|
| Program Director | 1 | 1 |
| Cohort/Accelerator Manager | 1 | 2 |
| Trade & Market Entry Lead | 1 | 2 |
| Platform/Data Analyst | 1 | 1 |
| Admin/Finance/M&E | 1 | 2 |
| **รวม** | **6** | **9** |

- **ทักษะวิกฤต:** trade/market entry, business matchmaking ตปท., บริหาร cohort accelerator, market intelligence
- **บทบาทหน่วยงาน:** DITP เจ้าภาพ (mission/pavilion/สคต.) · DAAT (curate+mentor) · AICAT (คัดกรองเทคโนโลยี/IP) · BOI (ลงทุน) · สสว. (co-funding) · EXIM (การเงินส่งออก)
- **ความเสี่ยงคน:** ทีมเล็กต้องประสานหลายหน่วย → ใช้เครือข่าย สคต. ของ DITP ที่มีอยู่ทั่วโลก

---

## สรุปกำลังคนรวม (ภาพรวมโครงการ)

| โครงการ | ทีม core (FTE) | Pool เสริม |
|---|---:|---|
| 1. Foundation Model | 16 | + annotate 25, advisor นานาชาติ |
| 2. Summit | 6 | + PCO จ้างเหมา, speaker/buyer |
| 3. Sandbox | 9 | + mentor pool |
| 4. Toolkit | 15–20 | + content/mentor pool, มหาวิทยาลัย |
| 5. Export | 9 | + สคต. เครือข่าย DITP |
| **รวม core** | **~55–60 FTE** | + pool/จ้างเหมา/in-kind จำนวนมาก |

- **ทักษะที่ขาดแคลน & ต้องวางแผนล่วงหน้า:** นักวิจัย AI อาวุโส (F1) · MLOps/FinOps (F1,F3) · instructional designer (F4) · business matchmaker ตปท. (F2,F5)
- **กลยุทธ์คน:** ทีมแกนรัฐ/AICAT เล็ก + **pool/จ้างเหมา/in-kind** (PCO, mentor, มหาวิทยาลัย, สคต.) + ดึงคนไทยในต่างประเทศ/advisor นานาชาติเฉพาะจุดวิกฤต

> ตัวเลขเป็นประมาณการเชิงวางแผน · งบ/อัตรากำลังจริงกำหนดหลัง Feasibility Study (Gate เดือน 6) — รายละเอียด unit rate + แหล่งอ้างอิงครบในไฟล์ 36
