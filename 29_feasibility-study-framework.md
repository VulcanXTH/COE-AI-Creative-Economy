# 29 — Feasibility Study Framework: 5 Flagship Projects

> **สถานะ:** ร่างกรอบการศึกษา (Phase 1, เดือน 1–6)
> **อ้างอิง:** 26_output-outcome-mega-v3.md · แนวทาง ดร.อภิรดี (Output เชิงคุณภาพ → KPI จริงหลัง Feasibility)
> **วันที่:** 4 มิ.ย. 2569

---

## 1. หลักการ

- Feasibility Study คือ **เงื่อนไขก่อนตั้ง KPI ตัวเลข** — ห้ามเสนอตัวเลขผูกมัดก่อนผลการศึกษาออก
- ทุก Flagship ต้องผ่าน **Go / No-Go / Pivot Gate** ภายในเดือนที่ 6
- ผลการศึกษาใช้เป็นเอกสารแนบ **คำของบประมาณปี 1 (83 MB)** ต่อ สดช. / สำนักงบประมาณ

## 2. Template กลาง (ใช้ทุก Flagship)

แต่ละโปรเจกต์ต้องตอบ 7 หมวด:

| หมวด | คำถามหลัก | หลักฐานที่ต้องมี |
|---|---|---|
| A. Demand | ใครคือผู้ใช้? ต้องการจริงไหม? | LOI (ดูไฟล์ 30) + สัมภาษณ์ ≥10 ราย/โปรเจกต์ |
| B. Technical | ทำได้จริงด้วยเทคโนโลยี/ทีมที่หาได้ไหม? | Technical assessment โดยผู้เชี่ยวชาญ ≥2 ราย |
| C. Financial | งบเท่าไหร่ (Min/Mid/Max)? แหล่งทุน? | Cost breakdown + เทียบ benchmark ต่างประเทศ |
| D. Organizational | ใครเป็นเจ้าภาพ? มีคนทำจริงไหม? | ตัว host ยืนยันเป็นลายลักษณ์อักษร |
| E. Legal/IP | ติดข้อกฎหมาย/ลิขสิทธิ์อะไร? | ความเห็นผู้ทรงคุณวุฒิกฎหมาย IP/AI |
| F. Risk | อะไรทำให้ล้มเหลว? mitigation? | Risk register ≥5 ข้อ/โปรเจกต์ |
| G. KPI Proposal | ถ้า Go — KPI ตัวเลขที่สมเหตุสมผลคืออะไร? | ตัวเลขอิงผล A–C ไม่ใช่ aspiration |

**เกณฑ์ Gate:**
- **Go** = ผ่าน A+B+C+D ครบ (Demand ชัด, ทำได้, งบมีทาง, มีเจ้าภาพ)
- **Pivot** = Demand ชัดแต่ scope ต้องเปลี่ยน → เสนอ scope ใหม่ใน 1 เดือน
- **No-Go** = Demand ไม่ชัด หรือไม่มีเจ้าภาพ → คืนงบเข้า pool ของ Flagship อื่น

## 3. รายโปรเจกต์

### 🏛 F1: Thai Creative AI Foundation Model

| ประเด็น | รายละเอียดที่ต้องศึกษา |
|---|---|
| คำถามชี้ขาด | Build / Fine-tune / Partner — ทางไหนคุ้มสุด? |
| Demand | สตูดิโอ/เอเจนซีไทยยอมเปลี่ยนจาก tool ต่างชาติไหม? เงื่อนไขคืออะไร (ราคา/ภาษาไทย/ลิขสิทธิ์)? |
| Technical | ฐาน Open Source ตัวเลือก (เทียบ ≥3 model) · ต้องการ GPU/Compute เท่าไหร่ · Thai Creative Dataset (Pillar 3) พร้อมแค่ไหน |
| Financial | สมมติฐานเริ่มต้น: Fine-tune 30–60 MB vs Build from scratch 300+ MB — ต้อง validate |
| เจ้าภาพศึกษา | **BDI + AICAT** + นักวิจัย AI (ที่ปรึกษา Annex B) |
| ความเสี่ยงหลัก | Model ต่างชาติพัฒนาเร็วกว่า → ต้องนิยาม "ช่องว่างที่ไทยชนะได้" (ภาษาไทย/วัฒนธรรม/ลิขสิทธิ์สะอาด) |
| Gate target | เดือนที่ 6 |

### 🎪 F2: Bangkok AI Creative Market

| ประเด็น | รายละเอียดที่ต้องศึกษา |
|---|---|
| คำถามชี้ขาด | จัด standalone หรือเกาะ event DITP เดิม (Bangkok International Content Market)? |
| Demand | ผู้ซื้อต่างชาติสนใจ AI Creative content ไทยจริงไหม? — สำรวจผ่านเครือข่าย DITP |
| Financial | สมมติฐาน: เกาะ event เดิมปี 1 (5–10 MB) → standalone ปี 3 (30+ MB) |
| เจ้าภาพศึกษา | **DITP + DAAT + AICAT** |
| ความเสี่ยงหลัก | จัดงานแล้วไม่มี deal เกิด → ต้องมี matchmaking pipeline ก่อนงาน |
| Gate target | เดือนที่ 4 (ต้องเร็ว — ถ้า Go ต้อง lock วันงานปี 2570) |

### 🧪 F3: National AI Creative Sandbox

| ประเด็น | รายละเอียดที่ต้องศึกษา |
|---|---|
| คำถามชี้ขาด | ใครเป็น platform host (BDI? depa? เอกชน?) และเกณฑ์คัดโปรเจกต์เป็นอย่างไร? |
| Demand | SME/ครีเอเตอร์กี่รายพร้อมเข้าทดสอบ? ปัญหาที่อยากแก้คืออะไร (ต้นทุน? เครื่องมือ? ข้อมูล?) |
| Technical | ต้องการ infra อะไร: compute credits, licensed tools, dataset access |
| Financial | งบ Pillar 4 = 35% ของงบรวม — ก้อนใหญ่สุด ต้อง breakdown ละเอียดสุด |
| เจ้าภาพศึกษา | **NIA + BDI + AICAT** |
| ความเสี่ยงหลัก | กลายเป็น "แจกของฟรี" ไม่เกิด case ขยายพาณิชย์ → เกณฑ์คัดต้องมี path to commercialization |
| Gate target | เดือนที่ 5 |

### 🎓 F4: Creator AI Toolkit (เชื่อม OFOS)

| ประเด็น | รายละเอียดที่ต้องศึกษา |
|---|---|
| คำถามชี้ขาด | เชื่อมระบบ OFOS (One Family One Soft Power) ทางเทคนิค+งบประมาณอย่างไร? ใครจ่ายค่า license? |
| Demand | สาย "AI ขั้นสูง" ใน OFOS มีผู้สมัครเป้าหมายกี่คน? skill ไหนตลาดต้องการ (โยง Talent Market Survey) |
| Organizational | ต้องประสาน **THACCA** — ปัจจุบันไม่มีในคณะ 17 ที่นั่ง (ช่องว่าง stakeholder ข้อ 4) |
| Financial | สมมติฐาน: Toolkit + คอร์ส 15–25 MB ปี 1 |
| เจ้าภาพศึกษา | **depa + CEA + AICAT** |
| ความเสี่ยงหลัก | OFOS เป็นโครงการการเมือง — ถ้านโยบายเปลี่ยน ต้องมี standalone mode |
| Gate target | เดือนที่ 6 |

### 🌐 F5: AI Creative Export Accelerator

| ประเด็น | รายละเอียดที่ต้องศึกษา |
|---|---|
| คำถามชี้ขาด | DITP รับเป็นเจ้าภาพร่วมอย่างเป็นทางการไหม? cohort แรกกี่ราย เกณฑ์อะไร? |
| Demand | บริษัท content/tool ไทยที่ "พร้อม export ถ้ามีคนช่วย" มีกี่ราย — สำรวจผ่าน DAAT/TACGA/สมาคมผู้กำกับ |
| Financial | สมมติฐาน: cohort 10–15 ราย ปีละ 10–15 MB |
| เจ้าภาพศึกษา | **DITP + DAAT** |
| ความเสี่ยงหลัก | ซ้ำซ้อนโครงการ export เดิมของ DITP → ต้องนิยาม unique value (AI-specific) |
| Gate target | เดือนที่ 5 |

## 4. Timeline รวม (เดือน 1–6)

```
เดือน 1   ตั้งทีมศึกษา 5 ชุด + อนุมัติ TOR โดยคณะทำงาน (ประชุมครั้งที่ 1)
เดือน 1–2 ส่ง LOI + แบบสำรวจ (ไฟล์ 30) ถึง 10+ องค์กร
เดือน 2–4 สัมภาษณ์เชิงลึก + Technical/Financial assessment
เดือน 4   Gate F2 (Bangkok AI Creative Market) — ตัดสินใจก่อนเพื่อ lock วันงาน
เดือน 5   Gate F3 (Sandbox) + F5 (Export Accelerator)
เดือน 6   Gate F1 (Foundation Model) + F4 (Toolkit/OFOS)
          → สรุปรายงาน Feasibility + ชุด KPI ตัวจริง เสนอคณะทำงาน (ประชุมครั้งที่ 3)
```

## 5. งบการศึกษา (จากงบปี 1)

| รายการ | ประมาณการ |
|---|---|
| ทีมศึกษา/ที่ปรึกษา 5 โปรเจกต์ | 3.0 MB |
| สำรวจ Demand + Talent Market Survey | 1.5 MB |
| Technical assessment (F1 เป็นหลัก) | 1.0 MB |
| ประชุม/workshop ผู้มีส่วนได้ส่วนเสีย | 0.5 MB |
| **รวม** | **6.0 MB** (~7% ของงบปี 1) |

## 6. ผู้รับผิดชอบ

- **เจ้าภาพรวม:** เลขานุการคณะทำงาน (AICAT) — รวบรวมรายงาน 5 ชุดเข้า Gate
- **ผู้กลั่นกรอง Gate:** อนุคณะ Sandbox & Data (NIA+BDI+AICAT) เสนอ → คณะทำงานหลักลงมติ
- **เชื่อมโยง:** ผล Demand จาก LOI (ไฟล์ 30) + baseline จาก Cost Index (ไฟล์ 31) ป้อนเข้าหมวด A และ G ของทุกโปรเจกต์
