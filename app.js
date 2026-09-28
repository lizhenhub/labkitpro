/* LabKitPro SPA — Third-Party Best-Practice content base
 * Four modules: 1 Generate · 2 Review · 3 Uncertainty · 4 Diagnose
 * English-first (LANG="en"); Chinese pack is a future domestic release.
 * State auto-persists to localStorage so overseas users can resume work.
 */
(function () {
  "use strict";
  var DATA = window.LABKIT_DATA;
  var LANG = "en";
  var BRAND = "LabKitPro";
  var CONTACT_EMAIL = "support@labkitpro.com";

  // === Phase 2 (Supabase lead capture) — GDPR EU data residency ===
  // P0: account/lead data MUST live in Supabase EU region (Frankfurt, eu-central-1).
  // Leave both blank until you supply the real project URL + anon key; captureLead() no-ops when SUPABASE_URL is empty.
  var SUPABASE_URL    = "https://katoiificgcdwiekxjye.supabase.co";   // eu-central-1 (Frankfurt) — user-provisioned
  var SUPABASE_ANON   = "sb_publishable_u7kOZrJtzhmx033NImhclQ_jk584-X5";   // anon PUBLIC key only — NEVER the service_role key
  var LEAD_TABLE      = "leads";
  // leads.email is NOT NULL, so exports by anonymous visitors are recorded under this sentinel
  // until they volunteer a real address (the modal below). Filter it out when mailing.
  var PLACEHOLDER_EMAIL = "uncollected@labkitpro.local";

  // Fire-and-forget lead capture. Runs only after cookie consent (GDPR) and when configured.
  function captureLead(action, docCode) {
    if (!SUPABASE_URL || !SUPABASE_ANON) return;          // not configured → silent no-op
    var c = cookieConsent();
    if (!c || !c.analytics) return;                       // GDPR: only with explicit consent
    try {
      var payload = {
        action: action,                                  // 'export' | 'export+email' | 'generate' | 'package'
        doc_code: docCode || null,
        region: (labInfo && labInfo.country) || null,
        lab_name: (labInfo && labInfo.labName) || null,
        email: (labInfo && labInfo.email) || PLACEHOLDER_EMAIL,
        hp: getHpValue()                                 // honeypot: always "" for real users; bots that fill it get rejected by RLS
      };
      fetch(SUPABASE_URL + "/rest/v1/" + LEAD_TABLE, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "apikey": SUPABASE_ANON,
          "Authorization": "Bearer " + SUPABASE_ANON,
          "Prefer": "return=minimal"
        },
        body: JSON.stringify(payload)
      }).catch(function () {});                          // best-effort; never block UX
    } catch (e) {}
  }

  // NOTE: the optional email-capture popup has been removed entirely — the site
  // never shows a modal. Exports are still recorded silently by captureLead()
  // above (only after cookie consent). getHpValue() stays so the payload shape
  // and the DB-side RLS honeypot gate remain valid; with no visible form it is
  // always "" for real users.
  function getHpValue() {
    var el = $("hp");
    return el ? (el.value || "").trim() : "";
  }
  var STORE_KEY = "labkitpro_v1";
  var sel = null;            // current generation doc
  var labInfo = {};          // shared lab profile (persisted)
  var audit = {};            // {itemIndex: 'pass'|'obs'|'fail'}
  var diag = {};             // {secId: {maturity, evidence}}
  var gumComps = [];         // GUM components (persisted)
  var mcmState = { comps: [], model: "a + b", N: 200000, seed: 20240927 };
  var diagRegion = "GLOBAL", auditRegion = "GLOBAL";
  var treeFilter = "";
  var lastUnc = null;        // {method, uc, U, detail, kinfo, html}

  /* ---------- bilingual UI ---------- */
  var T = {
    zh: {
      gen: "文件生成", audit: "合规审查", unc: "不确定度", diag: "现实诊断",
      home: "总览", homeEyebrow: "ISO/IEC 17025 · 全球通用最佳实践基准",
      homeTitle: "满足 ISO/IEC 17025，你需要的都在这里",
      homeIntro: "诊断差距、生成 87 份最佳实践文件、跑完整合规审查、计算测量不确定度——一个工具箱全搞定，无需顾问。",
      homeModules: "四大核心版块",
      cta1: "诊断我的实验室", cta2: "生成文件", cta3: "运行合规审查",
      homeOpen: "进入 →", homeBack: "返回总览",
      homeStats: "87 份文件 · 119 项审查 · GUM + MCM · §4–§8 差距热力图",
      appTitle: "LabKitPro · ISO/IEC 17025 实验室工具箱",
      appSub: "四大核心版块 — 现实诊断 · 文件生成 · 合规审查 · 不确定度",
      foot: "依据 ISO/IEC 17025:2017 ｜ 实验室管理全球通用最佳实践基准",
      treeTitle: "文件层级", levelQM: "① 质量手册", levelLP: "② 程序文件",
      levelWI: "③ 作业指导书", levelFR: "④ 表单记录", levelSG: "⑤ 专项指南",
      metaTitle: "文件信息", labTitle: "实验室信息（动态填充）",
      f_labName: "实验室名称", f_type: "实验室类型", f_field: "检测/校准领域",
      f_country: "国家 / 地区", f_body: "认可机构", f_param: "关键参数",
      typeOpts: ["检测实验室", "校准实验室", "二者兼具"],
      bodyOpts: ["UKAS（英国）", "A2LA（美国）", "NATA（澳大利亚）", "CNAS/SAC（中国）", "DAkkS（德国）", "其他 ILAC 成员"],
      apply: "套用实验室信息", copy: "复制 HTML", print: "打印 / 另存 PDF", export: "导出独立 HTML",
      genDoc: "生成文档", fieldsTitle: "必填字段占位（表单类）",
      auditIntro: "以 FR-25 内部审核检查表（119 项，§4–§8 全条款）为审查规则库；SG-01 / SG-02 提供内审 / 管评判据。逐项勾选判定后实时统计符合性。",
      auditReport: "导出审查报告", pass: "符合", obs: "观察", fail: "不符合", na: "未审",
      sgRef: "内审 / 管评判据（SG-01 / SG-02）",
      uncIntro: "基于 SG-03（GUM 总论）/ SG-04（7 例数值案例）/ SG-05（MCM 6 例，N=200000 固定种子）/ WI-08 构建计算引擎，支持 GUM 与 MCM 两种方法；FR-13 / FR-38 作为结果输出模板。",
      gum: "GUM 法（解析合成）", mcm: "MCM 法（蒙特卡洛，N=200000 固定种子）",
      cName: "分量名称", cType: "类型", cDist: "分布", cParams: "参数(a,b或μ,σ)", cSens: "灵敏度c",
      typeA: "A 类", typeB: "B 类", distN: "正态", distU: "均匀", distT: "三角",
      addComp: "＋ 增加分量", calc: "计算", loadGumPreset: "载入 GUM 示例", loadMcmPreset: "载入 MCM 示例",
      model: "模型 y =", mcmN: "抽样次数 N", seed: "随机种子(固定)", run: "运行 MCM",
      uc: "合成标准不确定度 u꜀", U: "扩展不确定度 U (k=2)", mean: "均值", std: "标准差",
      outTitle: "结果输出（FR-13 / FR-38 模板）", exportUnc: "导出不确定度报告",
      diagIntro: "基于 SG-06 条款骨架 + 88 份文件清单（条款↔文件映射）生成诊断要素热力图权重（§4–§8）。上传/填写证据并按现状成熟度评分，输出差距热力图。",
      diagHeat: "差距热力图（权重 × (1 − 成熟度/5)）", diagMap: "条款 ↔ 文件映射（诊断要素上下文）",
      maturity: "现状成熟度 (0–5)", evidence: "证据 / 说明", genHeat: "生成差距热力图",
      docsCount: "相关文件", selPlaceholder: "← 从左侧选择文件层级与具体文件", metaType: "文件类型", metaParent: "归属程序", metaClause: "对应 ISO 17025 条款", metaBasis: "编制依据", metaExtras: "附加依据", auditPointsLbl: "审核要点：", auditEvLbl: "证据与记录：", auditNoFr25: "FR-25 未找到", repTitle: "内部审核符合性报告", repBasis: "依据 ISO/IEC 17025:2017", repDate: "生成日期：", repItems: "项",       repConform: "符合", repObserve: "观察", repNonconf: "不符合", repFail: "不符合：",
      regionGlobal: "全球通用基准（ISO/IEC 17025 通用）", regionLbl: "区域 / 认可机构", regAddTitle: "区域附录", regFocusLbl: "机构特定重点", regNcLbl: "该机构常见不符合项", regNoteGlobal: "全球通用基准 —— 通用 ISO/IEC 17025:2017 要求。",
      auditConform: "审核符合 {p}%", fromAudit: "↺ 按审核重置", mismatch: "自评分与审核不符", gapPlan: "差距整改清单", weakest: "最弱条款", uncOnFile: "已测算不确定度", statusAudit: "审查", statusDiag: "诊断", statusUnc: "不确定度", autoFromAudit: "（按审核）", manualTag: "手动",
      guideTitle: "如何使用本文档", gWhat: "本文用途", gOwner: "负责角色", gWhen: "使用时机", gInput: "你需要提供", gOutput: "交付物", gClause: "对应条款", gRelated: "关联文档", gTodo: "待补全项", gFilled: "已自动套用实验室名称", gFillHint: "本文含约 {n} 处需你补全的空白与未填节（标记 — 或 ____），请按你实验室实际情况填写。", ownerQM: "最高管理者 + 质量负责人", ownerLP: "程序负责人 / 质量负责人", ownerWI: "技术主管 / 检测校准人员", ownerFR: "执行人 / 记录保管", ownerSG: "体系 / 技术负责人（参考）", whenQM: "体系建立、换版或外部评审前", whenLP: "日常运行、内部审核、管理评审时引用", whenWI: "执行具体方法或操作时", whenFR: "活动发生时填写并存档", whenSG: "作为方法 / 依据参考", outDoc: "可打印或导出的独立 .html，作为体系文件归档",
    },
    en: {
      gen: "Generate", audit: "Review", unc: "Uncertainty", diag: "Diagnose",
      home: "Overview", homeEyebrow: "ISO/IEC 17025 · Global best-practice baseline",
      homeTitle: "Everything your lab needs to meet ISO/IEC 17025",
      homeIntro: "Diagnose your gaps, generate 87 best-practice documents, run a full compliance review, and calculate measurement uncertainty — in one toolkit, no consultant required.",
      homeModules: "Four core modules",
      cta1: "Diagnose my lab", cta2: "Generate documents", cta3: "Run compliance review",
      homeOpen: "Open →", homeBack: "Back to overview",
      homeStats: "87 documents · 119 audit items · GUM + MCM · §4–§8 gap heatmap",
      appTitle: "LabKitPro · ISO/IEC 17025 Lab Toolkit",
      appSub: "Four core modules — Diagnose · Generate · Review · Uncertainty",
      foot: "Based on ISO/IEC 17025:2017 ｜ Global best-practice baseline for laboratory management",
      treeTitle: "Document library", levelQM: "① Quality Manual", levelLP: "② Procedures",
      levelWI: "③ Work Instructions", levelFR: "④ Forms / Records", levelSG: "⑤ Special Guides",
      metaTitle: "Document Info", labTitle: "Laboratory profile (auto-filled)",
      f_labName: "Lab name", f_type: "Lab type", f_field: "Field of testing/calibration",
      f_country: "Country / Region", f_body: "Accreditation body", f_param: "Key parameter",
      typeOpts: ["Testing lab", "Calibration lab", "Both"],
      bodyOpts: ["UKAS (UK)", "A2LA (US)", "NATA (AU)", "CNAS/SAC (CN)", "DAkkS (DE)", "Other ILAC member"],
      apply: "Apply lab profile", copy: "Copy HTML", print: "Print / PDF", export: "Export standalone HTML",
      genDoc: "Generate document", fieldsTitle: "Required fields (forms)",
      auditIntro: "FR-25 internal-audit checklist (119 items, §4–§8) is the review rule base; SG-01 / SG-02 give internal-audit / management-review criteria. Mark each judgment to get live conformance stats.",
      auditReport: "Export review report", pass: "Conform", obs: "Observe", fail: "Nonconform", na: "Not done",
      sgRef: "Internal-audit / mgmt-review criteria (SG-01 / SG-02)",
      uncIntro: "Engine built on SG-03 (GUM), SG-04 (7 numerical cases), SG-05 (MCM 6 cases, N=200000 fixed seed), WI-08; supports GUM & MCM. FR-13 / FR-38 as output templates.",
      gum: "GUM (analytical synthesis)", mcm: "MCM (Monte Carlo, N=200000 fixed seed)",
      cName: "Component", cType: "Type", cDist: "Distribution", cParams: "Params (a,b or μ,σ)", cSens: "Sensitivity c",
      typeA: "Type A", typeB: "Type B", distN: "Normal", distU: "Uniform", distT: "Triangular",
      addComp: "＋ Add component", calc: "Calculate", loadGumPreset: "Load GUM example", loadMcmPreset: "Load MCM example",
      model: "Model y =", mcmN: "Samples N", seed: "Seed (fixed)", run: "Run MCM",
      uc: "Combined std uncertainty u꜀", U: "Expanded uncertainty U (k=2)", mean: "Mean", std: "Std dev",
      outTitle: "Result output (FR-13 / FR-38 template)", exportUnc: "Export uncertainty report",
      diagIntro: "Diagnosis heatmap weights from SG-06 clause skeleton + 88-file clause↔file map (§4–§8). Rate your current maturity per clause to surface the gaps.",
      diagHeat: "Gap heatmap (weight × (1 − maturity/5))", diagMap: "Clause ↔ File map (diagnosis context)",
      maturity: "Current maturity (0–5)", evidence: "Evidence / notes", genHeat: "Generate gap heatmap",
      docsCount: "Related files", selPlaceholder: "← Select a level and document from the left", metaType: "Document type", metaParent: "Owning procedure", metaClause: "ISO/IEC 17025 clause", metaBasis: "Basis of preparation", metaExtras: "Additional basis", auditPointsLbl: "Audit points: ", auditEvLbl: "Evidence & records: ", auditNoFr25: "FR-25 not found", repTitle: "Internal Audit Conformance Report", repBasis: "Based on ISO/IEC 17025:2017", repDate: "Generated: ", repItems: "items",       repConform: "Conform", repObserve: "Observe", repNonconf: "Nonconform", repFail: "Nonconform: ",
      regionGlobal: "Global baseline (ISO/IEC 17025 generic)", regionLbl: "Region / Accreditation body", regAddTitle: "Regional addenda", regFocusLbl: "Body-specific focus", regNcLbl: "Common nonconformances reported by this body", regNoteGlobal: "Global baseline — generic ISO/IEC 17025:2017 requirements.",
      auditConform: "audit {p}%", fromAudit: "↺ reset to audit", mismatch: "self-rating differs from audit", gapPlan: "Your gap-closing plan", weakest: "weakest", uncOnFile: "Measurement uncertainty on file", statusAudit: "Review", statusDiag: "Diagnosis", statusUnc: "Uncertainty", autoFromAudit: "(from audit)", manualTag: "manual",
      guideTitle: "How to use this document", gWhat: "What this is", gOwner: "Owner / responsibility", gWhen: "When to use", gInput: "What you provide", gOutput: "Deliverable", gClause: "Clause", gRelated: "Related documents", gTodo: "To complete", gFilled: "Lab name auto-filled", gFillHint: "This document has about {n} blanks / unfilled sections (marked — or ____) for you to complete with your lab's actual details.", ownerQM: "Top Management + Quality Manager", ownerLP: "Process owner / Quality Manager", ownerWI: "Technical Supervisor / Testing-Calibration staff", ownerFR: "Operator / Records keeper", ownerSG: "System / Technical lead (reference)", whenQM: "At system build, revision, or before external assessment", whenLP: "Referenced in daily ops, internal audit, management review", whenWI: "When performing the method / operation", whenFR: "Fill when the activity occurs, then archive", whenSG: "As method / basis reference", outDoc: "Printable or exportable standalone .html, archived as a system document",
    }
  };
  function t(k) { return (T[LANG] && T[LANG][k]) || T.zh[k] || k; }
  function dtitle(d) { return LANG === "en" ? (d.code + " · " + d.name_en) : (d.code + " · " + d.name); }

  /* ---------- regional accreditation mapping appendix (P1) ----------
   * Curated, verified appendix: for each region the ISO/IEC 17025:2017 generic
   * clauses are overlaid with the specific accreditation body's additional
   * requirements and the nonconformances that body most often reports.
   * Sources (checked 2026-09): A2LA R/P-series; NVLAP NIST HB 150 + 15 CFR 285;
   * DAkkS R-17025-K / German Accreditation Body Act / Reg (EC) 765/2008;
   * UKAS LAB & M3003 series / Accreditation Regulations 2009;
   * NABL 100/112/162; SAC-SINGLAS 001 / SAC 01 / SAC 02. */
  var REGION_ORDER = ["US", "EU", "IN", "SG", "AU"];
  var REGIONS = {
    US: {
      name: "United States", bodies: "A2LA · NVLAP",
      intro: "US labs are accredited by A2LA (private, ILAC MRA signatory) and/or NVLAP (NIST-managed, governed by 15 CFR part 285). Both assess against ISO/IEC 17025:2017 plus body-specific criteria — A2LA R-series/P-series policies (R101 general, R102 conditions, R103 proficiency testing, R105 logo use, P102 traceability, P103 uncertainty) and NVLAP NIST Handbook 150 + program handbooks (e.g. HB 150-2 calibration).",
      sec: {
        "§4": { focus: "A2LA R101 requires a documented conflict-of-interest policy; P106/P120 govern branch & satellite sites. NVLAP requires named Authorized Representative & Approved Signatories and controls use of the 'NVLAP' name/logo (15 CFR 285).", nc: "Impartiality risk not documented for in-house/commercial labs; signatory authority ambiguous." },
        "§5": { focus: "A2LA P106 (branch systems) / P107 (transfer). NVLAP requires a cross-reference document mapping the lab's management system to ISO/IEC 17025 clauses 4–8 and HB 150 annexes whenever numbering differs.", nc: "Org chart does not reflect actual reporting lines; scope claimed broader than witnessed methods." },
        "§6": { focus: "A2LA P102 (instrument-specific metrological traceability) & P103 (uncertainty per discipline). NVLAP requires uncertainty budgets for ALL CMCs at application/renewal as source spreadsheets (not PDF-converted).", nc: "A2LA reports ~38% of its nonconformances under §6.4 (equipment cal records) and ~21% under §6.6 (externally provided products/services)." },
        "§7": { focus: "A2LA R103 makes proficiency testing mandatory at initial/renewal/surveillance/annual review (PT plan F104). NVLAP requires a PT participation plan at renewal. Methods must be validated per field.", nc: "PT does not cover every accredited field; test report omits date of issue (§7.8.2.1 j); method performed outside approved scope." },
        "§8": { focus: "A2LA R102 conditions; internal audit (§8.8) + management review (§8.9) records required. NVLAP expects a fully implemented QMS before application; on-site every other year, certificate valid 1 year.", nc: "Management review outputs missing required records (§8.9.3); internal audit records incomplete." }
      }
    },
    EU: {
      name: "Europe (EU / UK)", bodies: "DAkkS (DE) · UKAS (UK)",
      intro: "In Europe, DAkkS is the German national body (DIN EN ISO/IEC 17025:2018; R-17025-K for calibration; German Accreditation Body Act; Reg (EC) 765/2008) and UKAS is the UK national body (Accreditation Regulations 2009; extensive LAB/M3003 guidance). Both are EA MLA & ILAC MRA signatories.",
      sec: {
        "§4": { focus: "DAkkS assesses independence of the conformity assessment body. UKAS requires documented safeguards against commercial pressure and separation of calibration staff from production in in-house labs.", nc: "Impartiality not documented for embedded/in-house labs; customer confidentiality gaps." },
        "§5": { focus: "DAkkS: legal entity under German law; 'flexible accreditation' lets scope changes deploy without re-audit; docs submitted ~8 weeks before assessment. UKAS: the Schedule of Accreditation defines precise scope (parameter, range, best measurement capability).", nc: "Schedule of Accreditation does not match actual methods; branch-lab management system not integrated." },
        "§6": { focus: "DAkkS R-17025-K details §6; traceability to national standards (PTB). UKAS: traceability chain terminates at NPL; TPS 41 traceability policy; equipment-record evidence.", nc: "§6.4 equipment-record gaps; §6.5 traceability not unbroken; §6.2 personnel competence/authorisation records." },
        "§7": { focus: "DAkkS: method validation + uncertainty per R-17025; reporting per §7.8 with DAkkS cover sheet. UKAS: M3003 (uncertainty, GUM), LAB 5 (reporting), LAB 48 (decision rules), LAB 12 (testing uncertainty); decision rules mandatory.", nc: "§7.6 uncertainty budget gaps; §7.8 certificate content (expanded U, k, UKAS wording); missing decision-rule statements." },
        "§8": { focus: "DAkkS: 5-year reaccreditation, annual DAkkS review; QMS must mirror each 17025 clause. UKAS: 4-year re-assessment, annual surveillance; option A or B management system.", nc: "§8.8 internal audit & §8.9 management review records; §8.1 risk-based actions." }
      }
    },
    IN: {
      name: "India", bodies: "NABL",
      intro: "NABL (a constituent board of the Quality Council of India) is the sole Indian accreditation body and an ILAC MRA / APAC MLA signatory. It accredits to ISO/IEC 17025:2017 (testing & calibration), ISO 15189 (medical), ISO/IEC 17043 (PTP) and ISO 17034 (RMP), with NABL-specific criteria (NABL 100/112/162).",
      sec: {
        "§4": { focus: "NABL requires documented precautions against conflicts of interest and demonstrable impartiality — particularly for government / FSL labs.", nc: "Conflict-of-interest register missing for government-linked labs." },
        "§5": { focus: "Legal identity required; scope defined by analyte, matrix, technique, LOQ and method. NABL also runs a Temporary Site Laboratories scheme (aggregates/concrete) and G-LAP drinking-water program.", nc: "Scope claimed broader than witnessed techniques; org structure unclear." },
        "§6": { focus: "Traceability is mandatory; uncertainty must reflect real operating conditions (not ideal assumptions). Calibration certificates and uncertainty statements are scrutinised by assessors.", nc: "§6.2 personnel competence (selection/training/authorisation/monitoring); §6.4 equipment cal records; §6.5 traceability." },
        "§7": { focus: "Methods validated with data; PT participation expected (NABL 112 for forensics adds GeT-RM / CTS / NIJ schemes). Sampling treated as a technical activity; report content per §7.8.", nc: "§7.6 uncertainty not realistic; §7.7 PT not covering all fields; date of issue omitted from reports." },
        "§8": { focus: "Option A or B; initial accreditation 2 years, extended to 4 years on satisfactory surveillance (12–18 month intervals). NABL 100/112/162 provide guidance; NC closure typically 60 days (major) / 90 days (minor).", nc: "§8.8 internal audit; §8.9 management review outputs; delayed NC closure." }
      }
    },
    SG: {
      name: "Singapore", bodies: "SAC (SAC-SINGLAS)",
      intro: "SAC administers the SAC-SINGLAS scheme (managed by Enterprise Singapore), an APAC MLA & ILAC MRA signatory whose reports are recognised in 121 economies. Accreditation is to ISO/IEC 17025:2017 plus SAC-specific criteria (SAC-SINGLAS 001 process; SAC 01 terms; SAC 02 marks rules; discipline Technical Notes).",
      sec: {
        "§4": { focus: "Impartiality & independence per SAC 01 Terms & Conditions; confidentiality around use of the SAC accreditation / MRA marks.", nc: "Impartiality not documented for in-house labs; conflict declaration absent." },
        "§5": { focus: "Legally registered entity; accreditation granted only for specific tests/calibrations in defined fields; site & branch labs and a Management Representative are recognised.", nc: "Scope mismatch with application; branch lab not covered by the certificate." },
        "§6": { focus: "Traceability to national standards (A*STAR National Metrology Centre or an ILAC MRA lab); equipment master list with cal status; documented in-service checks between calibrations; personnel competency files.", nc: "§6.2 personnel authorisation records; §6.4 equipment cal status; §6.5 traceability to NMC." },
        "§7": { focus: "Validated method per parameter/range; documented measurement-uncertainty budget for every calibration; decision rules; reporting per §7.8; discipline-specific SAC Technical Notes apply.", nc: "§7.6 uncertainty budgets not updated; §7.8 certificate elements; §7.7 QC / PT coverage." },
        "§8": { focus: "Option A or B; SAC 02 governs use of accreditation & MRA marks; routine surveillance; typical 3–6 month path to accreditation.", nc: "§8.8 internal audit; §8.9 management review; document-control weaknesses." }
      }
    },
    AU: {
      name: "Australia", bodies: "NATA",
      intro: "NATA (National Association of Testing Authorities, Australia) is the sole Australian accreditation body and an ILAC MRA / APAC MLA signatory. It accredits to ISO/IEC 17025:2017 (testing & calibration) plus NATA-specific Accreditation Criteria — generic criteria spanning management & technical requirements and a library of field-specific technical criteria (e.g. chemistry, microbiology, construction materials, electrical) — and governs use of the NATA logo & MRA mark through the NATA Rules. NATA is recognised by the Australian Government and underpinned by the National Measurement Institute (NMI Australia) for metrological traceability.",
      sec: {
        "§4": { focus: "NATA criteria require documented impartiality & conflict-of-interest controls and demonstrable independence — especially for in-house, commercial and regulator-linked labs. Confidentiality around client information and the NATA mark is assessed.", nc: "Conflict-of-interest register absent for in-house / commercial labs; confidentiality gaps around client data." },
        "§5": { focus: "Legal entity & documented organisational structure; NATA issues a Scope of Accreditation listing methods, ranges and (for calibration) best measurement capability. Field-specific technical criteria must be met for each accredited discipline.", nc: "Scope of Accreditation does not match witnessed methods; org structure unclear; field-specific criteria not evidenced." },
        "§6": { focus: "Traceability to national standards via NMI Australia (or an ILAC MRA laboratory); documented measurement-uncertainty budgets; equipment master list with calibration status and in-service checks. NATA technical criteria specify equipment & environment per field.", nc: "§6.2 personnel competence / authorisation records; §6.4 equipment calibration records; §6.5 unbroken traceability to NMI." },
        "§7": { focus: "Methods validated with supporting data per field; documented uncertainty budgets for calibration; decision rules addressed per ILAC P14 / NATA guidance; reporting per §7.8; participation in proficiency testing expected for accredited fields.", nc: "§7.6 uncertainty not realistic; §7.7 PT not covering all fields; §7.8 certificate elements / date of issue omitted." },
        "§8": { focus: "Option A or B management system; NATA conducts initial assessment, periodic surveillance and reassessment; internal audit (§8.8) and management review (§8.9) records expected. Nonconformance & corrective-action closure tracked to NATA timeframes.", nc: "§8.8 internal audit; §8.9 management review outputs; delayed nonconformance closure." }
      }
    }
  };

  /* ---------- helpers ---------- */
  function $(id) { return document.getElementById(id); }
  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function el(html) { var d = document.createElement("div"); d.innerHTML = html.trim(); return d.firstChild; }
  function pick(d, k) { var v = (LANG === "en" && d[k + "_en"] != null && d[k + "_en"] !== "") ? d[k + "_en"] : d[k]; return v == null ? "" : v; }
  function opt(arr, sel) { return arr.map(function (x) { return "<option value='" + x[0] + "'" + (x[0] === sel ? " selected" : "") + ">" + esc(x[1]) + "</option>"; }).join(""); }
  function kv(k, v) { return "<dt>" + esc(k) + "</dt><dd>" + esc(v) + "</dd>"; }
  function fld(id, label, val) {
    return "<label class='fld' for='" + id + "'>" + esc(label) + "</label>" +
      "<input id='" + id + "' value='" + esc(val) + "'>";
  }
  function sel2(id, label, opts, val) {
    var o = opts.map(function (x) { return "<option" + (x === val ? " selected" : "") + ">" + esc(x) + "</option>"; }).join("");
    return "<label class='fld' for='" + id + "'>" + esc(label) + "</label><select id='" + id + "'>" + o + "</select>";
  }

  /* ---------- persistence ---------- */
  function saveState() {
    try {
      localStorage.setItem(STORE_KEY, JSON.stringify({
        labInfo: labInfo, audit: audit, diag: diag,
        gumComps: gumComps, mcmState: mcmState,
        diagRegion: diagRegion, auditRegion: auditRegion
      }));
    } catch (e) {}
  }
  function loadState() {
    try {
      var s = JSON.parse(localStorage.getItem(STORE_KEY) || "{}");
      if (s.labInfo) labInfo = s.labInfo;
      if (s.audit) audit = s.audit;
      if (s.diag) diag = s.diag;
      if (s.gumComps && s.gumComps.length) gumComps = s.gumComps;
      if (s.mcmState && s.mcmState.comps && s.mcmState.comps.length) mcmState = s.mcmState;
      if (s.diagRegion) diagRegion = s.diagRegion;
      if (s.auditRegion) auditRegion = s.auditRegion;
    } catch (e) {}
  }

  /* ---------- tabs ---------- */
  // Order = the four core modules: Diagnose → Generate → Review → Uncertainty.
  // "home" is the overview pane and carries no module number.
  var TABS = [
    { id: "home", pane: "pane-home", noNum: true },
    { id: "diag", pane: "pane-diag" },
    { id: "gen", pane: "pane-gen" },
    { id: "audit", pane: "pane-audit" },
    { id: "unc", pane: "pane-unc" }
  ];
  function renderTabs() {
    var box = $("tabs"); if (!box) return; box.innerHTML = "";
    var n = 0;
    TABS.forEach(function (tb) {
      var label = tb.noNum ? t("home") : t(tb.id);
      var num = tb.noNum ? "" : String(++n);
      var b = el('<button class="tab' + (tb.noNum ? " tab-home" : "") + '" data-id="' + tb.id + '">' +
        (num ? '<span class="num">' + num + "</span>" : "") + esc(label) + "</button>");
      b.onclick = function () { showPane(tb.id); };
      box.appendChild(b);
    });
  }
  function showPane(id) {
    if (!TABS.some(function (tb) { return tb.id === id; })) id = "home";
    TABS.forEach(function (tb) {
      var on = tb.id === id;
      document.getElementById(tb.pane).classList.toggle("active", on);
      var btn = document.querySelector('.tab[data-id="' + tb.id + '"]');
      if (btn) btn.classList.toggle("active", on);
    });
    if (id === "home") renderHome();
    if (id === "gen") renderTree();
    if (id === "audit") renderAudit();
    if (id === "unc") renderUnc();
    if (id === "diag") renderDiag();
  }

  /* ---------- 0 · Home — the four core modules ---------- */
  // Everything in the toolkit serves these four modules, so the overview leads
  // with them. Card copy lives here (zh + en) and follows the current LANG.
  var MODULES = [
    {
      id: "diag", n: "1", icon: "◎",
      zh: "现实诊断", en: "Diagnose",
      zhD: "对照 ISO/IEC 17025 §4–§8 评估实验室当前成熟度，生成差距热力图 —— 先看清短板，再决定投入。",
      enD: "Score your current maturity against ISO/IEC 17025 §4–§8 and get a gap heatmap — see the gaps before you invest."
    },
    {
      id: "gen", n: "2", icon: "▤",
      zh: "文件生成", en: "Generate",
      zhD: "实验室档案填一次，即可生成 87 份最佳实践文件：质量手册、程序文件、作业指导书、表单记录、专项指南。",
      enD: "Fill your lab profile once, then generate any of the 87 best-practice documents — manual, procedures, work instructions, forms and guides."
    },
    {
      id: "audit", n: "3", icon: "☑",
      zh: "合规审查", en: "Review",
      zhD: "用 FR-25 内部审核检查表（119 项，覆盖 §4–§8 全条款）逐项判定，实时统计符合性并导出审查报告。",
      enD: "Work the FR-25 internal-audit checklist (119 items, all of §4–§8), get live conformance stats and export a review report."
    },
    {
      id: "unc", n: "4", icon: "∑",
      zh: "不确定度", en: "Uncertainty",
      zhD: "用 GUM 解析法或蒙特卡洛（MCM，N=200000 固定种子）评定测量不确定度，输出 FR-13 / FR-38 报告模板。",
      enD: "Evaluate measurement uncertainty with GUM or Monte-Carlo (MCM, N=200000 fixed seed) and export an FR-13 / FR-38 report."
    }
  ];
  function renderHome() {
    var hero = $("homeHero"), grid = $("modGrid"), note = $("homeNote");
    if (hero) {
      hero.innerHTML =
        "<div class='home-hero'>" +
          "<span class='eyebrow'>" + esc(t("homeEyebrow")) + "</span>" +
          "<h2>" + esc(t("homeTitle")) + "</h2>" +
          "<p>" + esc(t("homeIntro")) + "</p>" +
          "<div class='home-cta'>" +
            "<button class='btn cta-p' data-goto='diag'>" + esc(t("cta1")) + "</button>" +
            "<button class='btn' data-goto='gen'>" + esc(t("cta2")) + "</button>" +
            "<button class='btn ghost' data-goto='audit'>" + esc(t("cta3")) + "</button>" +
          "</div>" +
        "</div>";
    }
    if (grid) {
      grid.innerHTML = "<div class='modhead'>" + esc(t("homeModules")) + "</div>";
      MODULES.forEach(function (m) {
        var c = el("<button class='modcard' data-mod='" + m.id + "'>" +
          "<div class='modcard-h'><span class='modnum'>" + m.n + "</span><span class='modicon'>" + m.icon + "</span></div>" +
          "<h3>" + esc(LANG === "en" ? m.en : m.zh) + "</h3>" +
          "<p>" + esc(LANG === "en" ? m.enD : m.zhD) + "</p>" +
          "<span class='modopen'>" + esc(t("homeOpen")) + "</span></button>");
        c.onclick = function () { showPane(m.id); };
        grid.appendChild(c);
      });
    }
    if (note) note.textContent = t("homeStats");
  }

  /* ================= 1 · Generate ================= */
  function byLevel(lv) { return DATA.docs.filter(function (d) { return d.level === lv; }); }
  function renderTree() {
    var tree = $("tree"); if (!tree) return; tree.innerHTML = "";
    tree.appendChild(el('<h4>' + t("treeTitle") + "</h4>"));
    renderFocus();
    var q = (treeFilter || "").toLowerCase();
    var groups = [["QM", t("levelQM"), "lvl-qm"], ["LP", t("levelLP"), "lvl-lp"],
    ["WI", t("levelWI"), "lvl-wi"], ["FR", t("levelFR"), "lvl-fr"], ["SG", t("levelSG"), "lvl-sg"]];
    var shown = 0;
    groups.forEach(function (g) {
      var list = byLevel(g[0]).sort(function (a, b) { return a.code < b.code ? -1 : 1; });
      list = list.filter(function (d) {
        if (!q) return true;
        return (d.code + " " + d.name_en + " " + d.name + " " + (d.clause || "")).toLowerCase().indexOf(q) >= 0;
      });
      if (!list.length) return;
      shown += list.length;
      var wrap = el('<div class="' + g[2] + '"></div>');
      wrap.appendChild(el("<h4>" + g[1] + " (" + list.length + ")</h4>"));
      var ch = el('<div class="children"></div>');
      list.forEach(function (d) {
        var n = el('<div class="node" data-code="' + d.code + '"><span>' + esc(d.code) +
          " · " + esc(LANG === "en" ? d.name_en : d.name) + "</span><span class='c'>" +
          (d.clause_sec || "") + "</span></div>");
        n.onclick = function () { selectDoc(d.code); };
        ch.appendChild(n);
      });
      wrap.appendChild(ch); tree.appendChild(wrap);
    });
    var hint = $("treeHint");
    if (hint) hint.textContent = shown + " document" + (shown === 1 ? "" : "s") + (q ? " match '" + treeFilter + "'" : " · search above");
  }
  function selectDoc(code) {
    sel = DATA.docs.filter(function (d) { return d.code === code; })[0];
    if (!sel) return;
    document.querySelectorAll(".tree .node").forEach(function (n) {
      n.classList.toggle("sel", n.getAttribute("data-code") === code);
    });
    renderGenMeta(); renderGenGuide(); renderLabForm(); renderGenButtons(); renderPreview();
  }
  function renderGenMeta() {
    var box = $("genMeta"); if (!sel) { box.innerHTML = ""; return; }
    var d = sel;
    var lvlClass = { QM: "qm", LP: "lp", WI: "wi", FR: "fr", SG: "sg" }[d.level];
    box.innerHTML =
      '<div class="card"><div style="display:flex;align-items:center;gap:10px">' +
      '<span class="badge ' + lvlClass + '">' + d.level + "</span>" +
      "<div><div style='font-size:16px;font-weight:800'>" + esc(dtitle(d)) + "</div>" +
      "<div class='en'>" + esc(d.name_en) + "</div></div></div><div class='kv' style='margin-top:12px'>" +
      kv(t("metaType"), pick(d, "type")) + kv(t("metaParent"), d.parent ? (d.parent + (d.parent_clause ? " ・ " + d.parent_clause : "")) : "—") +
      kv(t("metaClause"), d.clause || (d.parent_clause || "—")) +
      kv(t("metaBasis"), pick(d, "basis")) +
      (d.extras && d.extras.length ? kv(t("metaExtras"), d.extras.join(LANG === "en" ? ", " : "、")) : "") +
      "</div>" +
      "<div class='btnbar' style='margin-top:12px'>" +
      "<button class='btn ghost' id='bToAudit'>→ " + t("audit") + " (FR-25)</button>" +
      "<button class='btn ghost' id='bToDiag'>→ " + t("diag") + " (clause " + (d.clause_sec || d.clause || "—") + ")</button>" +
      "</div></div>";
    $("bToAudit").onclick = function () { showPane("audit"); };
    $("bToDiag").onclick = function () { showPane("diag"); };
  }

  /* ---------- ordinary-user usability: per-doc guidance + smart pre-fill ---------- */
  var GUIDE_OWNER = { QM: "ownerQM", LP: "ownerLP", WI: "ownerWI", FR: "ownerFR", SG: "ownerSG" };
  var GUIDE_WHEN = { QM: "whenQM", LP: "whenLP", WI: "whenWI", FR: "whenFR", SG: "whenSG" };

  function guideItem(k, v) { return "<div class='gi'><div class='gk'>" + esc(k) + "</div><div class='gv'>" + v + "</div></div>"; }

  function docGuidance(d) {
    if (!d) return "";
    var rel = {};
    if (d.parent) rel[d.parent] = 1;
    var refs = (pick(d, "body") || "").match(/data-doc='[^']+'/g) || [];
    refs.forEach(function (r) {
      var m = r.match(/data-doc='([^']+)'/);
      var c = m ? m[1] : "";
      if (c && c !== d.code && DATA.docs.some(function (x) { return x.code === c; })) rel[c] = 1;
    });
    var relList = Object.keys(rel);
    var be = pick(d, "body") || "";
    var blanks = (be.match(/_{4,}/g) || []).length;
    var em = (be.match(/—|–/g) || []).length;
    var todo = blanks + em;
    var html = "<div class='guide'><div class='guide-h'>" + t("guideTitle") + "</div><div class='guide-grid'>" +
      guideItem(t("gWhat"), esc((d.level_en || d.level_cn || d.level)) + " · " + esc(pick(d, "type") || "—")) +
      guideItem(t("gOwner"), t(GUIDE_OWNER[d.level] || "gOwner")) +
      guideItem(t("gWhen"), t(GUIDE_WHEN[d.level] || "gWhen")) +
      guideItem(t("gClause"), esc(d.clause || (d.parent_clause || "—"))) +
      guideItem(t("gOutput"), t("outDoc")) +
      (relList.length ? guideItem(t("gRelated"), relList.map(function (c) { return "<code>" + esc(c) + "</code>"; }).join(" ")) : "") +
      "</div>";
    if (todo > 0) html += "<div class='guide-todo'>" + tf("gFillHint", { n: todo }) + "</div>";
    if (labInfo.labName) html += "<div class='guide-fill'>✓ " + t("gFilled") + "</div>";
    html += "</div>";
    return html;
  }

  function renderGenGuide() {
    var box = $("genGuide");
    if (!box) return;
    box.innerHTML = docGuidance(sel);
    box.style.display = sel ? "block" : "none";
  }

  function smartFill(html) {
    if (!labInfo.labName) return html;
    var name = esc(labInfo.labName);
    return html
      .replace(/\bthe laboratory\b/gi, "the " + name)
      .replace(/\bthe lab\b/gi, "the " + name);
  }

  function renderLabForm() {
    var box = $("labForm"); if (!sel) { box.innerHTML = ""; return; }
    box.innerHTML = "<h4 style='margin:0 0 6px'>" + t("labTitle") + "</h4>" +
      fld("labName", t("f_labName"), labInfo.labName || "") +
      "<div class='row2'>" +
      sel2("type", t("f_type"), t("typeOpts"), labInfo.type) +
      sel2("body", t("f_body"), t("bodyOpts"), labInfo.body) + "</div>" +
      fld("field", t("f_field"), labInfo.field || "") +
      "<div class='row2'>" + fld("country", t("f_country"), labInfo.country || "") +
      fld("param", t("f_param"), labInfo.param || "") + "</div>" +
      "<div class='note'>Your profile is saved automatically on this device and reused across all four modules.</div>";
    box.querySelectorAll("input,select").forEach(function (i) {
      i.oninput = function () {
        labInfo[i.id] = i.value; saveState(); renderProfile();
        if (sel && (i.id === "labName" || i.id === "type" || i.id === "field" || i.id === "body")) renderPreview();
      };
    });
  }
  function renderGenButtons() {
    var box = $("genBtns"); if (!sel) { box.innerHTML = ""; return; }
    box.innerHTML =
      "<button class='btn' id='bApply'>" + t("apply") + "</button>" +
      "<button class='btn ghost' id='bCopy'>" + t("copy") + "</button>" +
      "<button class='btn ghost' id='bPrint'>" + t("print") + "</button>" +
      "<button class='btn ghost' id='bExport'>" + t("export") + "</button>";
    $("bApply").onclick = applyLab;
    $("bCopy").onclick = function () { copyText($("preview").innerHTML); };
    $("bPrint").onclick = function () { window.print(); };
    $("bExport").onclick = exportHTML;
  }
  function labBand() {
    if (!labInfo.labName && !labInfo.type) return "";
    return "<div class='labband'><b>" + esc(labInfo.labName || "Lab Name") + "</b> ｜ " +
      esc(labInfo.type || "—") + " ｜ " + esc(labInfo.field || "—") + " ｜ " +
      esc(labInfo.country || "—") + " ｜ " + esc(labInfo.body || "—") +
      (labInfo.param ? " ｜ " + esc(labInfo.param) : "") + "</div>";
  }
  function renderPreview() {
    var pv = $("preview"); if (!sel) { pv.innerHTML = "<p class='muted'>" + t("selPlaceholder") + "</p>"; return; }
    var html = smartFill(pick(sel, "body") || "");
    var extra = "";
    var flds = pick(sel, "fields");
    if (sel.level === "FR" && flds && flds.length && !pick(sel, "checklist").length) {
      extra = "<h3>" + t("fieldsTitle") + "</h3>";
      flds.forEach(function (f, i) {
        extra += "<div class='fillrow'><span class='lab'>" + esc(f.label) + "</span>" +
          "<input id='ff_" + i + "' value='" + esc(f.placeholder) + "'></div>";
      });
    }
    pv.innerHTML = labBand() + extra + html;
    pv.setAttribute("contenteditable", "true");
  }
  function applyLab() { renderPreview(); flash(t("apply")); }
  function flash(msg) {
    var b = $("bApply"); if (!b) return; var o = b.textContent; b.textContent = "✓ " + msg;
    setTimeout(function () { b.textContent = o; }, 1200);
  }
  function copyText(s) {
    var ta = document.createElement("textarea"); ta.value = s; document.body.appendChild(ta);
    ta.select(); try { document.execCommand("copy"); } catch (e) {} ta.remove();
    flash(t("copy"));
  }
  function exportHTML() {
    if (!sel) return;
    // capture any filled FR form fields
    var frFilled = "";
    if (sel.level === "FR") {
      var rows = $("preview").querySelectorAll(".fillrow");
      if (rows.length) {
        frFilled = "<h3>" + esc(t("fieldsTitle")) + "</h3><table class='comp-table'><tr><th>Field</th><th>Value</th></tr>";
        rows.forEach(function (r) {
          var lab = r.querySelector(".lab").textContent;
          var val = r.querySelector("input").value;
          frFilled += "<tr><td>" + esc(lab) + "</td><td>" + esc(val) + "</td></tr>";
        });
        frFilled += "</table>";
      }
    }
      // strip in-app cross-reference anchors to plain text for the standalone deliverable
    var bodyHtml = smartFill((pick(sel, "body") || "").replace(/<a\s[^>]*data-doc=[^>]*>([\s\S]*?)<\/a>/g, "$1"));
    var html = "<!DOCTYPE html><html lang='en'><head><meta charset='UTF-8'>" +
      "<title>" + esc(sel.code + " " + pick(sel, "name")) + "</title><style>" +
      "body{font-family:-apple-system,'Segoe UI',sans-serif;line-height:1.6;color:#1B2B36;max-width:900px;margin:0 auto;padding:24px}" +
      "h2{font-size:18px;border-bottom:2px solid #D8E2EA;padding-bottom:6px;margin:16px 0 8px}" +
      "table{border-collapse:collapse;width:100%;margin:10px 0}th{background:#0B2530;color:#fff;padding:7px 9px;text-align:left}" +
      "td{border:1px solid #D8E2EA;padding:6px 9px}ul,ol{margin:7px 0 7px 22px}</style></head><body>" +
      labBand() + frFilled + bodyHtml +
      "<hr style='margin:18px 0'><div style='font-size:11px;color:#5B6E7C;text-align:center'>Generated by " + BRAND + " · Based on ISO/IEC 17025:2017 (global generic baseline) · Adapt and revise this document to your actual circumstances before use.</div>" +
      "</body></html>";
    download(sel.code + "_" + (LANG === "en" ? sel.name_en : sel.name) + ".html", html);
    captureLead("export", sel.code);           // Phase 2: record which doc a user exports
  }
  function download(name, content) {
    var b = new Blob([content], { type: "text/html;charset=utf-8" });
    var a = document.createElement("a"); a.href = URL.createObjectURL(b); a.download = name; a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
  }

  /* ================= 2 · Review ================= */
  function renderAudit() {
    $("auditIntro").innerHTML = "<div class='callout info'>" + esc(t("auditIntro")) + "</div>";
    $("auditBtns").innerHTML = "<button class='btn' id='bRep'>" + t("auditReport") + "</button>" +
      "<button class='btn ghost' id='bPack'>⤓ " + t("gen") + " readiness package</button>";
    $("bRep").onclick = exportAuditReport;
    $("bPack").onclick = exportPackage;
    $("auditRegionBox").innerHTML = regionSelector("audit");
    wireRegion("audit");
    var fr25 = DATA.docs.filter(function (d) { return d.code === "FR-25"; })[0];
    var body = $("auditBody"); body.innerHTML = "";
    if (!fr25 || !pick(fr25, "checklist").length) { body.innerHTML = "<p class='muted'>" + t("auditNoFr25") + "</p>"; return; }
    var items = pick(fr25, "checklist");
    var secOrder = ["§4", "§5", "§6", "§7", "§8"];
    var groups = {};
    items.forEach(function (it, idx) {
      var sec = secOrder.filter(function (x) { return (it.clause || "").indexOf(x) === 0; })[0] || "§8";
      (groups[sec] = groups[sec] || []).push({ it: it, idx: idx });
    });
    secOrder.forEach(function (sec) {
      if (!groups[sec]) return;
      var g = el("<div class='clause-grp'><div class='hd2'><span>" + sec +
        " " + secLabel(sec) + "</span><span class='tally' data-sec='" + sec + "'></span></div></div>");
      groups[sec].forEach(function (o) {
        var it = o.it;
        var secId = (it.clause || "").replace(/[\d.]/g, "").replace("§", "§") || "§8";
        var row = el("<div class='chk'><div class='top'><span class='cl'>" + esc(it.clause) +
          "</span><span class='it'>" + esc(it.item) + "</span></div>" +
          "<div class='pt'>" + t("auditPointsLbl") + esc(it.points) + "</div>" +
          "<div class='ev'>" + t("auditEvLbl") + esc(it.evidence) + "</div>" +
          "<div class='judge'>" +
          jlabel("pass", o.idx, t("pass")) + jlabel("obs", o.idx, t("obs")) + jlabel("fail", o.idx, t("fail")) +
          "</div>" +
          "<div class='fixlink'><a data-goto='diag'>↳ diagnose / generate docs for " + esc(secId) + " →</a></div>" +
          "</div>");
        g.appendChild(row);
      });
      body.appendChild(g);
    });
    // SG refs
    var sgref = el("<div class='card'><h4 style='margin:0 0 8px'>" + t("sgRef") + "</h4></div>");
    ["SG-01", "SG-02"].forEach(function (c) {
      var d = DATA.docs.filter(function (x) { return x.code === c; })[0];
      if (d) {
        var det = el("<details style='margin:6px 0'><summary style='cursor:pointer;font-weight:700'>" +
          esc(dtitle(d)) + "</summary><div class='preview' style='margin-top:8px'>" + (pick(d, "body") || "") + "</div></details>");
        sgref.appendChild(det);
      }
    });
    body.appendChild(sgref);
    // restore prior judgments
    document.querySelectorAll(".judge input").forEach(function (inp) {
      var idx = parseInt(inp.name.replace("j", ""), 10);
      if (audit[idx] === inp.value) inp.checked = true;
    });
    // cross-link: uncertainty result as §6 evidence
    if (lastUnc && lastUnc.U != null) {
      var uo = el("<div class='card unc-onfile'><h4 style='margin:0 0 6px'>" + t("uncOnFile") + "</h4>" +
        "<div class='evid'>" + esc(t("statusUnc")) + ": U = " + esc(String(lastUnc.U)) + (lastUnc.method ? " · " + esc(lastUnc.method) : "") + "</div>" +
        "<div class='muted'>" + t("auditEvLbl") + "§6 measurement uncertainty — attached as objective evidence for the review.</div></div>");
      body.appendChild(uo);
    }
    updateAuditTally();
  }
  function jlabel(cls, idx, label) {
    return "<label class='" + cls + "'><input type='radio' name='j" + idx + "' value='" + cls + "' onchange='LKB.onJudge(" + idx + ",\"" + cls + "\")'><span>" + esc(label) + "</span></label>";
  }
  window.LKB = {
    onJudge: function (idx, v) { audit[idx] = v; updateAuditTally(); saveState(); if (currentPane() === "diag") genHeat(); if (currentPane() === "gen") renderFocus(); updateStatusBar(); },
    onMaturity: function (sec, v) { diag[sec] = diag[sec] || {}; diag[sec].maturity = parseInt(v); diag[sec].manual = true; $("mv_" + sec).textContent = v + "/5"; saveState(); if (currentPane() === "gen") renderFocus(); updateStatusBar(); },
    onEvidence: function (sec, v) { diag[sec] = diag[sec] || {}; diag[sec].evidence = v; saveState(); },
    resetMaturity: function (sec) { if (diag[sec]) { delete diag[sec].manual; if (diag[sec].maturity != null && !Object.keys(diag[sec]).filter(function (k) { return k !== "evidence"; }).length) delete diag[sec]; } saveState(); renderDiag(); genHeat(); updateStatusBar(); }
  };
  function secLabel(s) {
    var m = DATA.clause_sections.filter(function (x) { return x.id === s; })[0]; return m ? (LANG === "en" ? m.en : m.cn) : "";
  }
  function updateAuditTally() {
    var pAll = 0, oAll = 0, fAll = 0, nAll = 0;
    document.querySelectorAll(".tally[data-sec]").forEach(function (td) {
      var sec = td.getAttribute("data-sec");
      var p = 0, o = 0, f = 0, n = 0;
      pick(DATA.docs.filter(function (d) { return d.code === "FR-25"; })[0], "checklist").forEach(function (it, idx) {
        if ((it.clause || "").indexOf(sec) !== 0) return; n++;
        var v = audit[idx]; if (v === "pass") { p++; pAll++; } else if (v === "obs") { o++; oAll++; } else if (v === "fail") { f++; fAll++; }
        nAll++;
      });
      td.innerHTML = "<span style='color:#047857'>✓" + p + "</span> <span style='color:#B45309'>○" + o + "</span> <span style='color:#B91C1C'>✗" + f + "</span>";
    });
    var sum = $("auditSummary");
    if (sum) {
      var done = pAll + oAll + fAll;
      var pct = done ? Math.round(pAll / done * 100) : 0;
      sum.style.display = done ? "block" : "none";
      sum.innerHTML = "<b>Overall conformance:</b> " + pct + "% (" + pAll + " conform · " + oAll + " observe · " + fAll +
        " nonconform of " + done + " judged / " + nAll + " total) " +
        (fAll ? "<span class='muted'>— " + fAll + " open nonconformity" + (fAll === 1 ? "" : "ies") + "; jump to Diagnose to draft fixes.</span>" : "<span class='muted'>— no open nonconformities recorded.</span>");
    }
  }
  /* ---- cross-section coupling helpers ---- */
  function fr25Items() { var fr = DATA.docs.filter(function (d) { return d.code === "FR-25"; })[0]; return fr ? pick(fr, "checklist") : []; }
  function auditConformance(sec) {
    var items = fr25Items(), p = 0, o = 0, f = 0, n = 0;
    items.forEach(function (it, idx) {
      if ((it.clause || "").indexOf(sec) !== 0) return; n++;
      var v = audit[idx]; if (v === "pass") { p++; } else if (v === "obs") { o++; } else if (v === "fail") { f++; }
    });
    return { p: p, o: o, f: f, n: n, pct: n ? Math.round((p + o * 0.5) / n * 100) : 0 };
  }
  function auditMaturity(sec) {
    var c = auditConformance(sec); if (!c.n) return 0;
    if (c.f === c.n) return 0;                 // all nonconform
    if (c.p === c.n) return 5;                 // all conform
    var r = c.pct;
    return r >= 80 ? 4 : r >= 60 ? 3 : r >= 40 ? 2 : r > 0 ? 1 : 0;
  }
  function effMaturity(sec) {
    var d = diag[sec];
    if (d && d.manual) return d.maturity || 0;
    return auditMaturity(sec);
  }
  function tf(k, obj) { var s = t(k); if (obj) s = s.replace(/\{(\w+)\}/g, function (_, k2) { return obj[k2] != null ? obj[k2] : "{" + k2 + "}"; }); return s; }
  function updateStatusBar() {
    var bar = $("statusBar"); if (!bar) return;
    var c = auditConformance("§4"); var judged = fr25Items().filter(function (_, i) { return audit[i]; }).length;
    var total = fr25Items().length;
    var secs = DATA.clause_sections, gaps = {};
    secs.forEach(function (s) { var m = effMaturity(s.id); gaps[s.id] = s.weight * (1 - m / 5); });
    var rated = secs.filter(function (s) { return effMaturity(s.id) > 0 || auditConformance(s.id).n; });
    var overall = Math.round((1 - secs.reduce(function (a, s) { return a + gaps[s.id]; }, 0) / secs.reduce(function (a, s) { return a + s.weight; }, 0)) * 100);
    var u = lastUnc ? ("U=" + (lastUnc.U != null ? lastUnc.U : "—")) : "—";
    bar.innerHTML = "<span class='st'><b>" + t("statusAudit") + "</b> " + judged + "/" + total + "</span>" +
      "<span class='st'><b>" + t("statusDiag") + "</b> " + overall + "%</span>" +
      "<span class='st'><b>" + t("statusUnc") + "</b> " + u + "</span>";
    ["audit", "diag", "gen", "unc"].forEach(function (id) { var tb = document.querySelector(".tab[data-id='" + id + "']"); if (tb) tb.classList.toggle("has-data", id === "audit" ? judged > 0 : id === "diag" ? rated.length > 0 : id === "unc" ? !!lastUnc : false); });
  }
  function exportAuditReport() {
    var fr25 = DATA.docs.filter(function (d) { return d.code === "FR-25"; })[0];
    var lines = [BRAND + " — " + t("repTitle"), t("repBasis"), t("repDate") + new Date().toISOString().slice(0, 10), ""];
    if (labInfo.labName) lines.push("Laboratory: " + labInfo.labName + (labInfo.body ? " (" + labInfo.body + ")" : ""));
    lines.push("");
    var secOrder = ["§4", "§5", "§6", "§7", "§8"];
    secOrder.forEach(function (sec) {
      var p = 0, o = 0, f = 0, n = 0, fails = [];
      pick(fr25, "checklist").forEach(function (it, idx) {
        if ((it.clause || "").indexOf(sec) !== 0) return; n++;
        var v = audit[idx]; if (v === "pass") p++; else if (v === "obs") o++; else if (v === "fail") { f++; fails.push(it.clause + " " + it.item); }
      });
      lines.push(sec + " " + secLabel(sec) + ": " + n + " " + t("repItems") + " | " + t("repConform") + " " + p + " | " + t("repObserve") + " " + o + " | " + t("repNonconf") + " " + f);
      fails.forEach(function (x) { lines.push("  - " + t("repFail") + x); });
      if (sec === "§6" && lastUnc && lastUnc.U != null) lines.push("  - " + t("uncOnFile") + ": U = " + lastUnc.U + (lastUnc.method ? " (" + lastUnc.method + ")" : ""));
    });
    if (auditRegion !== "GLOBAL" && REGIONS[auditRegion]) {
      var R = REGIONS[auditRegion];
      lines.push("", "=== " + t("regAddTitle") + ": " + R.name + " (" + R.bodies + ") ===");
      DATA.clause_sections.forEach(function (s) {
        var e = R.sec[s.id]; if (!e) return;
        lines.push(s.id + " " + (LANG === "en" ? s.en : s.cn) + ":");
        lines.push("  - " + t("regFocusLbl") + " " + e.focus);
        lines.push("  - " + t("regNcLbl") + " " + e.nc);
      });
    }
    download("FR-25_Audit_Report.txt", lines.join("\n"));
    captureLead("export-audit", "FR-25");
  }

  /* ================= 3 · Uncertainty ================= */
  function renderUnc() {
    $("uncIntro").innerHTML = "<div class='callout info'>" + esc(t("uncIntro")) + "</div>";
    $("gumCard").innerHTML = "<h4 style='margin:0 0 8px'>" + t("gum") + "</h4>" +
      "<div id='gumTable'></div>" +
      "<button class='btn ghost' id='bGumAdd'>" + t("addComp") + "</button> " +
      "<button class='btn' id='bGumCalc'>" + t("calc") + "</button> " +
      "<button class='btn ghost' id='bGumPre'>" + t("loadGumPreset") + "</button>" +
      "<div id='gumRes' style='margin-top:10px'></div>";
    $("mcmCard").innerHTML = "<h4 style='margin:0 0 8px'>" + t("mcm") + "</h4>" +
      "<label class='fld'>" + t("model") + "</label><input id='mcmModel' value='" + mcmState.model + "'>" +
      "<div id='mcmComps'></div><button class='btn ghost' id='bMcmAdd'>" + t("addComp") + "</button> " +
      "<button class='btn ghost' id='bMcmPre'>" + t("loadMcmPreset") + "</button><br>" +
      "<label class='fld'>" + t("mcmN") + "</label><input id='mcmN' value='" + mcmState.N + "'>" +
      "<label class='fld'>" + t("seed") + "</label><input id='mcmSeed' value='" + mcmState.seed + "'>" +
      "<div class='btnbar'><button class='btn' id='bMcmRun'>" + t("run") + "</button></div>" +
      "<canvas id='mcmHist' class='hist'></canvas><div id='mcmRes'></div>";
    $("uncOut").innerHTML = "<h4 style='margin:0'>" + t("outTitle") + "</h4><div id='uncOutBody' class='muted'>" +
      (lastUnc ? lastUnc.html : t("selPlaceholder")) + "</div>" +
      "<div class='btnbar'>" +
      "<button class='btn ghost' id='bUncExp'>" + t("exportUnc") + "</button>" +
      (lastUnc ? "<button class='btn ghost' id='bUncOpen'>→ Open FR-13 template</button>" : "") + "</div>";
    $("bGumAdd").onclick = function () { gumComps.push({ name: "u" + (gumComps.length + 1), type: "B", dist: "U", a: 1, b: 1, c: 1 }); renderGumTable(); saveState(); };
    $("bGumCalc").onclick = calcGum;
    $("bGumPre").onclick = loadGumPreset;
    $("bMcmAdd").onclick = function () { mcmState.comps.push({ name: "x" + (mcmState.comps.length + 1), dist: "N", mu: 0, sigma: 1 }); renderMcmComps(); saveState(); };
    $("bMcmPre").onclick = loadMcmPreset;
    $("bMcmRun").onclick = runMcm;
    $("bUncExp").onclick = exportUnc;
    var bo = $("bUncOpen"); if (bo) bo.onclick = function () { showPane("gen"); selectDoc("FR-13"); };
    renderGumTable(); renderMcmComps();
    if (!lastUnc) calcGum(); // show a default GUM result on first open
  }
  function renderGumTable() {
    var h = "<table class='comp-table'><tr><th>" + t("cName") + "</th><th>" + t("cType") + "</th><th>" + t("cDist") +
      "</th><th>" + t("cParams") + " (a,b)</th><th>" + t("cSens") + "</th></tr>";
    gumComps.forEach(function (c, i) {
      h += "<tr><td><input id='g_n_" + i + "' value='" + esc(c.name) + "' size='6'></td>" +
        "<td><select id='g_t_" + i + "'>" + opt([["A", t("typeA")], ["B", t("typeB")]], c.type) + "</select></td>" +
        "<td><select id='g_d_" + i + "'>" + opt([["N", t("distN")], ["U", t("distU")], ["T", t("distT")]], c.dist) + "</select></td>" +
        "<td><input id='g_a_" + i + "' value='" + c.a + "' size='6'> , <input id='g_b_" + i + "' value='" + c.b + "' size='6'></td>" +
        "<td><input id='g_c_" + i + "' value='" + c.c + "' size='5'></td></tr>";
    });
    h += "</table>";
    var gt = $("gumTable"); if (!gt) return; gt.innerHTML = h;
    gumComps.forEach(function (c, i) {
      ["n", "t", "d", "a", "b", "c"].forEach(function (k) {
        var e = $("g_" + k + "_" + i); if (e) e.oninput = function () { readGum(); saveState(); };
      });
    });
  }
  function readGum() {
    gumComps.forEach(function (c, i) {
      c.name = $("g_n_" + i).value; c.type = $("g_t_" + i).value; c.dist = $("g_d_" + i).value;
      c.a = parseFloat($("g_a_" + i).value) || 0; c.b = parseFloat($("g_b_" + i).value) || 0; c.c = parseFloat($("g_c_" + i).value) || 1;
    });
  }
  function calcGum() {
    readGum();
    var rows = []; var sum = 0;
    gumComps.forEach(function (c) {
      var u; // standard uncertainty of this component
      if (c.dist === "N") u = c.b;                 // σ given
      else if (c.dist === "U") u = c.b / Math.sqrt(3); // half-width a=b
      else if (c.dist === "T") u = c.b / Math.sqrt(6);
      else u = c.b;
      var contrib = (c.c || 1) * u;
      sum += contrib * contrib;
      rows.push({ name: c.name, u: u, c: c.c, contrib: contrib });
    });
    var uc = Math.sqrt(sum); var U = 2 * uc;
    var tbl = "<table class='comp-table'><tr><th>" + t("cName") + "</th><th>uᵢ</th><th>cᵢ</th><th>cᵢ·uᵢ</th></tr>";
    rows.forEach(function (r) { tbl += "<tr><td>" + esc(r.name) + "</td><td>" + r.u.toFixed(5) + "</td><td>" + (r.c || 1) + "</td><td>" + r.contrib.toFixed(5) + "</td></tr>"; });
    tbl += "</table>";
    $("gumRes").innerHTML = tbl + "<div class='result'>" + t("uc") + " = <b>" + uc.toFixed(5) + "</b> ｜ " + t("U") + " = <b>" + U.toFixed(5) + "</b> (k=2)</div>";
    uncOut("GUM", uc, U, rows.map(function (r) { return r.name + ": u=" + r.u.toFixed(5); }).join("; "), "k=2");
  }
  function defaultGumComps() {
    return [
      { name: "u₁ Repeatability", type: "A", dist: "N", a: 0, b: 0.020, c: 1 },
      { name: "u₂ Reference standard", type: "B", dist: "N", a: 0, b: 0.015, c: 1 },
      { name: "u₃ Ice point", type: "B", dist: "U", a: 0, b: 0.010, c: 1 },
      { name: "u₄ Resolution", type: "B", dist: "U", a: 0, b: 0.029, c: 1 },
      { name: "u₅ Temperature fluctuation", type: "B", dist: "U", a: 0, b: 0.005, c: 1 }
    ];
  }
  function loadGumPreset() {
    gumComps = defaultGumComps();
    renderGumTable(); calcGum(); saveState();
  }
  function renderMcmComps() {
    var h = "";
    mcmState.comps.forEach(function (c, i) {
      h += "<div class='fillrow'><span class='lab'><input id='m_n_" + i + "' value='" + esc(c.name) + "' size='4'></span>" +
        "<span><select id='m_d_" + i + "'>" + opt([["N", t("distN")], ["U", t("distU")], ["T", t("distT")]], c.dist) + "</select> " +
        "μ<input id='m_mu_" + i + "' value='" + c.mu + "' size='5'> σ/a<input id='m_s_" + i + "' value='" + c.sigma + "' size='5'></span></div>";
    });
    var mc = $("mcmComps"); if (!mc) return; mc.innerHTML = h;
    mcmState.comps.forEach(function (c, i) {
      ["n", "d", "mu", "s"].forEach(function (k) { var e = $("m_" + k + "_" + i); if (e) e.oninput = function () { readMcm(); saveState(); }; });
    });
  }
  function readMcm() {
    mcmState.comps.forEach(function (c, i) {
      c.name = $("m_n_" + i).value; c.dist = $("m_d_" + i).value;
      c.mu = parseFloat($("m_mu_" + i).value) || 0; c.sigma = parseFloat($("m_s_" + i).value) || 1;
    });
    mcmState.model = $("mcmModel").value; mcmState.N = parseInt($("mcmN").value) || 200000; mcmState.seed = parseInt($("mcmSeed").value) || 20240927;
    saveState();
  }
  function mulberry32(a) { return function () { a |= 0; a = a + 0x6D2B79F5 | 0; var t = Math.imul(a ^ a >>> 15, 1 | a); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; }; }
  function gauss(rng) { var u = 0, v = 0; while (!u) u = rng(); while (!v) v = rng(); return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v); }
  function sample(dist, mu, sigma, rng) {
    if (dist === "N") return mu + sigma * gauss(rng);
    if (dist === "U") return mu + sigma * (rng() * 2 - 1);
    if (dist === "T") { var r = rng() * 2 - 1; return mu + sigma * r; }
    return mu;
  }
  function runMcm() {
    readMcm();
    var N = Math.min(mcmState.N, 500000); var rng = mulberry32(mcmState.seed >>> 0);
    var names = mcmState.comps.map(function (c) { return c.name.replace(/[^A-Za-z0-9_]/g, ""); });
    var fn; try { fn = new Function(names.join(","), "return " + mcmState.model + ";"); } catch (e) { $("mcmRes").innerHTML = "<div class='callout warn'>Invalid model syntax</div>"; return; }
    var ys = new Float64Array(N); var bad = 0;
    for (var i = 0; i < N; i++) {
      var args = mcmState.comps.map(function (c) { return sample(c.dist, c.mu, c.sigma, rng); });
      var y; try { y = fn.apply(null, args); } catch (e) { y = NaN; }
      if (isFinite(y)) ys[i] = y; else bad++;
    }
    var arr = Array.from(ys); arr.sort(function (a, b) { return a - b; });
    var mean = arr.reduce(function (s, x) { return s + x; }, 0) / arr.length;
    var variance = arr.reduce(function (s, x) { return s + (x - mean) * (x - mean); }, 0) / arr.length;
    var std = Math.sqrt(variance);
    var uc = std; var U = 2 * std;
    function pct(p) { return arr[Math.min(arr.length - 1, Math.floor(p * arr.length))]; }
    var lo = pct(0.025), hi = pct(0.975);
    $("mcmRes").innerHTML = "<div class='result'>" + t("mean") + " = <b>" + mean.toFixed(5) + "</b> ｜ " +
      t("std") + " = <b>" + std.toFixed(5) + "</b><br>" + t("uc") + " = <b>" + uc.toFixed(5) + "</b> ｜ " +
      t("U") + " = <b>" + U.toFixed(5) + "</b> (k=2)<br>95% interval [" + lo.toFixed(5) + ", " + hi.toFixed(5) + "]" +
      (bad ? "<br><span class='muted'>Invalid samples: " + bad + "</span>" : "") + "</div>";
    drawHist(arr, mean, std);
    uncOut("MCM (N=" + N + ")", uc, U, "Mean=" + mean.toFixed(5) + "; 95% interval=[" + lo.toFixed(5) + "," + hi.toFixed(5) + "]", "k=2, fixed seed " + mcmState.seed);
  }
  function drawHist(arr, mean, std) {
    var cv = $("mcmHist"); if (!cv) return; var ctx = cv.getContext("2d");
    var W = cv.width = cv.clientWidth || 600, H = cv.height = 160;
    ctx.clearRect(0, 0, W, H);
    var bins = 40, min = arr[0], max = arr[arr.length - 1], span = (max - min) || 1;
    var hist = new Array(bins).fill(0);
    arr.forEach(function (x) { var b = Math.min(bins - 1, Math.floor((x - min) / span * bins)); hist[b]++; });
    var mx = Math.max.apply(null, hist);
    ctx.fillStyle = "#0E7C9B";
    for (var i = 0; i < bins; i++) {
      var h = hist[i] / mx * (H - 20);
      ctx.fillRect(i * W / bins, H - h, W / bins - 1, h);
    }
    ctx.strokeStyle = "#B91C1C"; ctx.beginPath();
    var mx2 = (mean - min) / span * W; ctx.moveTo(mx2, 0); ctx.lineTo(mx2, H); ctx.stroke();
  }
  function defaultMcmComps() {
    return [{ name: "a", dist: "N", mu: 100, sigma: 0.05 }, { name: "b", dist: "N", mu: 0, sigma: 0.03 }];
  }
  function loadMcmPreset() {
    mcmState.model = "a + b"; mcmState.comps = defaultMcmComps();
    renderMcmComps(); saveState();
  }
  function uncOut(method, uc, U, detail, kinfo) {
    var d = DATA.docs.filter(function (x) { return x.code === "FR-13"; })[0];
    var html = "<div class='card' style='margin-top:12px'>";
    if (d) html += "<div class='muted'>" + esc(d.code) + " " + esc(LANG === "en" ? d.name_en : d.name) + " template</div>";
    html += "<div class='result'>" + t("uc") + " = <b>" + uc.toFixed(5) + "</b> ｜ " + t("U") + " = <b>" + U.toFixed(5) + "</b> (" + (kinfo || "k=2") + ")</div>" +
      "<p class='note'>Method: " + esc(method) + " ｜ " + esc(detail) + "</p>" +
      "<p class='note'>Based on ISO/IEC 17025:2017 §7.6 · JCGM 100/101 (+ILAC P14). Re-estimate with your laboratory's actual data.</p></div>";
    lastUnc = { method: method, uc: uc, U: U, detail: detail, kinfo: kinfo, html: html };
    var ub = $("uncOutBody"); if (ub) { ub.innerHTML = html; ub.classList.remove("muted"); }
    if (currentPane() === "audit") renderAudit();
    updateStatusBar();
    saveState();
  }
  function exportUnc() {
    var body = lastUnc ? lastUnc.html : "";
    if (!body) { flash(t("exportUnc")); return; }
    download("Uncertainty_Report_FR-13.html", "<!DOCTYPE html><html lang='en'><head><meta charset='UTF-8'><title>Uncertainty Report</title><style>body{font-family:-apple-system,'Segoe UI',sans-serif;max-width:820px;margin:0 auto;padding:24px}.result{background:#EAF3F8;border:1px solid #B9D6E4;border-radius:10px;padding:12px;margin:10px 0}</style></head><body><h2>Measurement Uncertainty Report</h2>" + (labBand()) + body + "<hr><div style='font-size:11px;color:#5B6E7C;text-align:center'>Generated by " + BRAND + " · Based on ISO/IEC 17025:2017 (global baseline)</div></body></html>");
  }

  /* ================= regional addenda (P1) ================= */
  function regionSelector(kind) {
    var cur = kind === "diag" ? diagRegion : auditRegion;
    var opts = "<option value='GLOBAL'" + (cur === "GLOBAL" ? " selected" : "") + ">" + esc(t("regionGlobal")) + "</option>" +
      REGION_ORDER.map(function (r) {
        return "<option value='" + r + "'" + (cur === r ? " selected" : "") + ">" + esc(REGIONS[r].name) + " — " + esc(REGIONS[r].bodies) + "</option>";
      }).join("");
    return "<label class='fld'>" + esc(t("regionLbl")) + "</label>" +
      "<select id='" + (kind === "diag" ? "diagRegionSel" : "auditRegionSel") + "'>" + opts + "</select>";
  }
  function wireRegion(kind) {
    var sel = $(kind === "diag" ? "diagRegionSel" : "auditRegionSel");
    if (!sel) return;
    sel.onchange = function () {
      if (kind === "diag") diagRegion = this.value; else auditRegion = this.value;
      saveState(); renderRegionAddenda(kind);
    };
    renderRegionAddenda(kind);
  }
  function renderRegionAddenda(kind) {
    var cur = kind === "diag" ? diagRegion : auditRegion;
    var box = $(kind === "diag" ? "diagRegionAdd" : "auditRegionAdd");
    if (!box) return;
    if (cur === "GLOBAL" || !REGIONS[cur]) {
      box.innerHTML = "<div class='callout info'>" + esc(t("regNoteGlobal")) + "</div>";
      return;
    }
    var R = REGIONS[cur];
    var html = "<div class='card regcard'><h4 style='margin:0 0 6px'>" + esc(t("regAddTitle")) +
      " · " + esc(R.name) + " <span class='muted'>(" + esc(R.bodies) + ")</span></h4>" +
      "<p class='note'>" + esc(R.intro) + "</p>";
    DATA.clause_sections.forEach(function (s) {
      var e = R.sec[s.id]; if (!e) return;
      html += "<div class='regsec'><div class='hd3'>" + s.id + " " + (LANG === "en" ? s.en : s.cn) + "</div>" +
        "<div><b>" + esc(t("regFocusLbl")) + "</b> " + esc(e.focus) + "</div>" +
        "<div><b>" + esc(t("regNcLbl")) + "</b> " + esc(e.nc) + "</div></div>";
    });
    html += "<div class='cta-mini'>Need this tailored to your lab, or another body (AU/NATA, CN/CNAS, …)? " +
      "<a href='mailto:" + CONTACT_EMAIL + "?subject=" + encodeURIComponent(BRAND + " — accreditation-body addendum request") + "'>Request a tailored addendum →</a></div>";
    html += "</div>";
    box.innerHTML = html;
  }

  /* ================= 4 · Diagnose ================= */
  function renderDiag() {
    $("diagIntro").innerHTML = "<div class='callout info'>" + esc(t("diagIntro")) + "</div>";
    $("diagRegionBox").innerHTML = regionSelector("diag");
    wireRegion("diag");
    var secs = DATA.clause_sections;
    var html = "<h4 style='margin:0 0 8px'>" + t("diagHeat") + "</h4>";
    html += "<div class='heat' id='heatGrid'>";
    secs.forEach(function (s) {
      html += "<div class='cell' data-sec='" + s.id + "' style='background:#E8F4FB'><div>" +
        (LANG === "en" ? s.en : s.cn) + "</div><div class='v' id='hv_" + s.id + "'>—</div><div class='w'>w=" + s.weight + "</div></div>";
    });
    html += "</div>";
    html += "<div class='card'>" + secs.map(function (s) {
      var d = diag[s.id] || {};
      var c = auditConformance(s.id);
      var autoTag = d.manual ? "" : (c.n ? " <span class='tag auto'>" + t("autoFromAudit") + "</span>" : "");
      return "<div style='margin:8px 0'><b>" + s.id + " " + (LANG === "en" ? s.en : s.cn) + "</b> (weight w=" + s.weight + ")" + autoTag + "<br>" +
        "<label class='fld'>" + t("maturity") + "</label>" +
        "<input type='range' min='0' max='5' value='" + ((diag[s.id] && diag[s.id].maturity) || 0) + "' id='m_" + s.id + "' oninput='LKB.onMaturity(\"" + s.id + "\",this.value)'>" +
        "<span id='mv_" + s.id + "'>" + ((diag[s.id] && diag[s.id].maturity) || 0) + "/5</span>" +
        (c.n ? " <span class='muted'>" + tf("auditConform", { p: c.pct }) + "</span> <a class='mini' onclick='LKB.resetMaturity(\"" + s.id + "\")'>" + t("fromAudit") + "</a>" : "") +
        "<label class='fld'>" + t("evidence") + "</label>" +
        "<textarea id='e_" + s.id + "' rows='2' oninput='LKB.onEvidence(\"" + s.id + "\",this.value)'>" + esc(diag[s.id] ? diag[s.id].evidence || "" : "") + "</textarea>" +
        "<div class='fixlink'><a data-doc='" + clauseTopDoc(s.id) + "'>↳ generate / review documents for " + s.id + " →</a></div></div>";
    }).join("<hr class='sep'>") + "</div>";
    html += "<div class='btnbar'><button class='btn' id='bHeat'>" + t("genHeat") + "</button>" +
      "<button class='btn ghost' id='bDiagExp'>⤓ " + t("gen") + " readiness package</button></div>";
    html += "<div id='diagSummary' class='card' style='display:none'></div>";
    $("diagHeat").innerHTML = html;
    $("bHeat").onclick = genHeat;
    $("bDiagExp").onclick = exportPackage;
    // mapping
    var map = "<h4 style='margin:0 0 8px'>" + t("diagMap") + "</h4><div class='mapgrid'>";
    secs.forEach(function (s) {
      var docs = DATA.docs.filter(function (d) { return d.clause_sec === s.id; });
      map += "<div class='mapcell'><div class='sec'>" + s.id + " " + (LANG === "en" ? s.en : s.cn) + "</div>" +
        "<div class='cnt'>" + t("docsCount") + ": " + docs.length + "</div>" +
        "<div style='font-size:11px;margin-top:4px'>" + docs.slice(0, 8).map(function (d) {
          return "<a class='codelink' data-doc='" + d.code + "' title='Open in Generator'>" + esc(d.code) + "</a>";
        }).join(" ") + (docs.length > 8 ? " …" : "") + "</div></div>";
    });
    map += "</div>";
    $("diagMap").innerHTML = map;
    genHeat();
    updateStatusBar();
  }
  function renderFocus() {
    var box = $("genFocus"); if (!box) return;
    var secs = DATA.clause_sections.slice().sort(function (a, b) { return effMaturity(a.id) - effMaturity(b.id); });
    var weak = secs.filter(function (s) { return effMaturity(s.id) < 5; }).slice(0, 2);
    if (!weak.length) { box.style.display = "none"; box.innerHTML = ""; return; }
    var html = "<h4 style='margin:0 0 8px'>" + t("gapPlan") + " <span class='muted'>(" + t("weakest") + ")</span></h4>";
    weak.forEach(function (s) {
      var docs = DATA.docs.filter(function (d) { return d.clause_sec === s.id; }).slice(0, 6);
      html += "<div class='focussec'><div class='fsh'><b>" + s.id + " " + (LANG === "en" ? s.en : s.cn) + "</b>" +
        " <span class='muted'>" + effMaturity(s.id) + "/5" + (diag[s.id] && diag[s.id].manual ? " " + t("manualTag") : "") + "</span></div>" +
        "<div class='fsdocs'>" + docs.map(function (d) {
          return "<a class='codelink' data-doc='" + d.code + "' title='Open in Generator'>" + esc(d.code) + "</a>";
        }).join(" ") + "</div></div>";
    });
    box.style.display = "block";
    box.innerHTML = html;
  }
  function clauseTopDoc(secId) {
    // first doc mapped to this clause, else first doc
    var d = DATA.docs.filter(function (x) { return x.clause_sec === secId; })[0];
    return d ? d.code : (DATA.docs[0] && DATA.docs[0].code);
  }
  function genHeat() {
    var secs = DATA.clause_sections;
    var gaps = {};
    secs.forEach(function (s) {
      var m = effMaturity(s.id);
      var gap = s.weight * (1 - m / 5);
      gaps[s.id] = gap;
      var cell = document.querySelector(".cell[data-sec='" + s.id + "']");
      if (cell) {
        cell.style.background = gapColor(gap);
        var c = auditConformance(s.id);
        var mm = (diag[s.id] && diag[s.id].manual && diag[s.id].maturity != null) ? diag[s.id].maturity : null;
        var note = c.n ? "<div class='au'>" + tf("auditConform", { p: c.pct }) + "</div>" : "";
        if (mm != null && c.n && Math.abs(mm - auditMaturity(s.id)) >= 2) note += "<div class='au warn'>" + t("mismatch") + "</div>";
        $("hv_" + s.id).innerHTML = (gap * 100).toFixed(0) + "%" + note;
      }
    });
    var rated = secs.filter(function (s) { return effMaturity(s.id) > 0 || auditConformance(s.id).n; });
    var overall = rated.length ? Math.round((1 - secs.reduce(function (a, s) { return a + gaps[s.id]; }, 0) / secs.reduce(function (a, s) { return a + s.weight; }, 0)) * 100) : 0;
    var sum = $("diagSummary");
    if (sum) {
      sum.style.display = "block";
      sum.innerHTML = "<b>Overall readiness:</b> " + overall + "% " +
        (rated.length < secs.length ? "<span class='muted'>(" + rated.length + "/" + secs.length + " clauses rated — rate all to finalise)</span>" : "<span class='muted'>— all clauses rated.</span>") +
        " <a data-goto='gen' style='margin-left:8px'>↳ generate the weakest documents →</a>";
    }
    saveState();
    if (currentPane() === "gen") renderFocus();
    updateStatusBar();
  }
  function gapColor(g) {
    var r, g2, b;
    if (g < 0.5) { var t = g / 0.5; r = Math.round(4 + t * 198); g2 = Math.round(120 + t * 60); b = Math.round(70 - t * 20); }
    else { var t2 = (g - 0.5) / 0.5; r = Math.round(202 + t2 * 0); g2 = Math.round(180 - t2 * 100); b = Math.round(50 - t2 * 20); }
    return "rgb(" + r + "," + g2 + "," + b + ")";
  }

  /* ================= readiness package (stickiness) ================= */
  function exportPackage() {
    var secs = DATA.clause_sections;
    var now = new Date().toISOString().slice(0, 10);
    var H = "<!DOCTYPE html><html lang='en'><head><meta charset='UTF-8'><title>" + BRAND + " Readiness Package</title>" +
      "<style>body{font-family:-apple-system,'Segoe UI',sans-serif;line-height:1.6;color:#1B2B36;max-width:920px;margin:0 auto;padding:24px}" +
      "h2{border-bottom:2px solid #0B2530;padding-bottom:6px;margin-top:26px}h3{color:#123448}.card{border:1px solid #D8E2EA;border-radius:10px;padding:12px 14px;margin:10px 0}" +
      "table{border-collapse:collapse;width:100%}td,th{border:1px solid #D8E2EA;padding:6px 9px;text-align:left}.muted{color:#5B6E7C}</style></head><body>";
    H += "<h1>" + BRAND + " — Laboratory Readiness Package</h1>";
    H += "<p class='muted'>Generated " + now + " · Based on ISO/IEC 17025:2017 (global generic baseline)</p>";
    H += labBand();
    // A. Diagnosis
    H += "<h2>1 · Diagnosis — gap heatmap</h2>";
    H += "<table><tr><th>Clause</th><th>Maturity</th><th>Gap</th><th>Evidence</th></tr>";
    var gaps = {};
    secs.forEach(function (s) {
      var m = (diag[s.id] && diag[s.id].maturity) || 0;
      var gap = s.weight * (1 - m / 5); gaps[s.id] = gap;
      H += "<tr><td>" + s.id + " " + (LANG === "en" ? s.en : s.cn) + "</td><td>" + m + "/5</td><td>" + (gap * 100).toFixed(0) + "%</td><td>" + esc(diag[s.id] ? diag[s.id].evidence || "" : "") + "</td></tr>";
    });
    H += "</table>";
    // B. Review
    H += "<h2>2 · Compliance review (FR-25)</h2>";
    var fr25 = DATA.docs.filter(function (d) { return d.code === "FR-25"; })[0];
    if (fr25) {
      var pAll = 0, oAll = 0, fAll = 0, nAll = 0; var fails = [];
      pick(fr25, "checklist").forEach(function (it, idx) {
        nAll++; var v = audit[idx];
        if (v === "pass") pAll++; else if (v === "obs") oAll++; else if (v === "fail") { fAll++; fails.push(it.clause + " " + it.item); }
      });
      var done = pAll + oAll + fAll;
      H += "<div class='card'>" + (done ? (Math.round(pAll / done * 100) + "% conform · " + pAll + " conform / " + oAll + " observe / " + fAll + " nonconform of " + done + " judged") : "No items judged yet.") + "</div>";
      if (fails.length) { H += "<p><b>Open nonconformities:</b></p><ul>" + fails.map(function (x) { return "<li>" + esc(x) + "</li>"; }).join("") + "</ul>"; }
    }
    // C. Uncertainty
    H += "<h2>3 · Uncertainty (last result)</h2>";
    H += lastUnc ? lastUnc.html : "<p class='muted'>No uncertainty calculation performed yet.</p>";
    // D. CTA
    H += "<hr><p style='font-size:12px;color:#5B6E7C'>Generated by " + BRAND + ". This package is a baseline self-assessment — for a laboratory-specific management system, contact " +
      "<a href='mailto:" + CONTACT_EMAIL + "'>" + CONTACT_EMAIL + "</a>.</p></body></html>";
    download(BRAND + "_Readiness_Package_" + now + ".html", H);
    captureLead("export-package", null);
  }

  /* ---------- lab profile chip ---------- */
  function renderProfile() {
    var chip = $("profileChip"); if (!chip) return;
    var name = labInfo.labName || "";
    chip.innerHTML = (name ? "<b>" + esc(name) + "</b>" : "<span class='muted'>No lab profile</span>") +
      (labInfo.body ? " · " + esc(labInfo.body) : "") +
      " <button class='mini' id='editProfile'>Edit profile</button>";
    var e = $("editProfile"); if (e) e.onclick = function () { showPane("gen"); var i = $("labName"); if (i) i.focus(); };
  }

  /* ---------- language toggle ---------- */
  function toggleLang() {
    var b = $("langBtn"); if (!b) return; var o = b.textContent; b.textContent = "即将上线"; setTimeout(function () { b.textContent = o; }, 1200);
  }
  function currentPane() { var a = document.querySelector(".pane.active"); if (!a) return "home"; var m = TABS.filter(function (x) { return x.pane === a.id; })[0]; return m ? m.id : "home"; }

  /* ---------- init ---------- */
  // ---- Cookie consent (GDPR) ----
  var COOKIE_KEY = "lk_cookie_consent";
  function cookieConsent() {
    try { return JSON.parse(localStorage.getItem(COOKIE_KEY) || "null"); } catch (e) { return null; }
  }
  function setCookieConsent(v) {
    localStorage.setItem(COOKIE_KEY, JSON.stringify({ analytics: v, ts: Date.now() }));
    if (v) window.dispatchEvent(new Event("lk-consent-granted")); // non-essential scripts may now load
    var b = $("cookieBanner"); if (b) b.hidden = true;
  }
  // gate for non-essential third-party scripts (analytics / live chat): call before injecting them
  function lkCanLoadNonEssential() { var c = cookieConsent(); return !!(c && c.analytics); }

  function init() {
    if (!DATA) { $("loading").textContent = "data.js missing"; return; }
    loadState();
    // seed default uncertainty components if nothing persisted (no DOM ops at init)
    if (!gumComps.length) gumComps = defaultGumComps();
    if (!mcmState.comps.length) mcmState.comps = defaultMcmComps();
    renderProfile();
    // cookie consent: show banner only if the user has not chosen yet
    if (!cookieConsent()) { var cb = $("cookieBanner"); if (cb) cb.hidden = false; }
    var ca = $("cookieAccept"); if (ca) ca.onclick = function () { setCookieConsent(true); };
    var cr = $("cookieReject"); if (cr) cr.onclick = function () { setCookieConsent(false); };
    // click the brand (title) to return to the module overview
    var at = $("appTitle");
    if (at) {
      at.onclick = function () { showPane("home"); };
      at.addEventListener("keydown", function (ev) {
        if (ev.key === "Enter" || ev.key === " ") { ev.preventDefault(); showPane("home"); }
      });
    }
    var cta = $("ctaContact"); if (cta) cta.href = "mailto:" + CONTACT_EMAIL + "?subject=" + encodeURIComponent(BRAND + " — tailored laboratory management system");
    $("loading").style.display = "none";
    $("langBtn").onclick = toggleLang;
    // search wiring
    var ts = $("treeSearch"); if (ts) ts.oninput = function () { treeFilter = this.value; renderTree(); };
    // in-app navigation: doc cross-refs + pane jumps
    document.addEventListener("click", function (e) {
      var a = e.target && e.target.closest ? e.target.closest("a[data-doc]") : null;
      if (a) { e.preventDefault(); var c = a.getAttribute("data-doc"); showPane("gen"); selectDoc(c); return; }
      var g = e.target && e.target.closest ? e.target.closest("[data-goto]") : null;
      if (g) { e.preventDefault(); showPane(g.getAttribute("data-goto")); return; }
    });
    renderTabs();
    showPane("home");
    updateStatusBar();
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init); else init();
})();
