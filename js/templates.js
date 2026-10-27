/* Docs: starter templates (our own text). Each returns plain document data that goes through DM.cleanDoc. */
(function () {
  "use strict";
  var M = window.DM;
  function p(style, text, props) { return M.para(style, text, props); }
  function r(t, m) { return M.withText(m || {}, t); }
  function li(text, k, l, props) { return M.para("normal", text, Object.assign({ list: { k: k || "ul", l: l || 0 } }, props || {})); }
  var today = function () { return new Date().toLocaleDateString(undefined, { year: "numeric", month: "long", day: "numeric" }); };

  var T = [
    { id: "blank", name: "Blank document", desc: "A4, Calibri 11", make: function () { return { title: "", blocks: [p("normal")] }; } },
    {
      id: "letter", name: "Letter", desc: "Formal block-style letter", make: function () {
        return {
          title: "Letter", page: { size: "a4", mt: 25.4, mb: 25.4, ml: 25.4, mr: 25.4 },
          blocks: [
            p("normal", [r("Jordan Lee", { b: true, z: 16, c: "#1d3557" })], { sa: 2 }),
            p("normal", "12 Garden Street, Springfield\njordan.lee@mail.example · +1 555 0134", { sa: 24, lh: 1.15 }),
            p("normal", today(), { sa: 18 }),
            p("normal", "Ms. Taylor Rivera\nHiring Manager\nNorthwind Studio\n200 Harbour Road, Springfield", { sa: 18 }),
            p("normal", "Re: Application for the Operations Coordinator role", { sa: 12 }),
            p("normal", "Dear Ms. Rivera,"),
            p("normal", "I am writing to apply for the Operations Coordinator position advertised on your website. Over the past four years I have kept a busy office running smoothly: planning schedules, managing suppliers and making sure the people around me have what they need to do their best work.", { align: "justify" }),
            p("normal", "In my current role I introduced a simple booking process that cut double-bookings to zero and saved the team several hours each week. I enjoy turning messy processes into calm, reliable ones, and I would bring the same care to Northwind Studio.", { align: "justify" }),
            p("normal", "Thank you for considering my application. I would welcome the chance to talk about how I could help, and I am available at your convenience.", { align: "justify" }),
            p("normal", "Sincerely,", { sb: 6, sa: 36 }),
            p("normal", "Jordan Lee")
          ]
        };
      }
    },
    {
      id: "cv", name: "CV / Résumé", desc: "One page, clean sections", make: function () {
        var t = M.table(2, 2, [["Tools", "Figma, Excel, Notion-style wikis, basic SQL"], ["Languages", "English (native), Arabic (fluent), French (conversational)"]]);
        t.hdr = false; t.cw = [0.25, 0.75];
        t.rows.forEach(function (row) { row[0].text = M.setMark(row[0].text, 0, M.len(row[0].text), "b", true); row[0].bg = "#f1f4f8"; });
        return {
          title: "CV", page: { size: "a4", mt: 16, mb: 16, ml: 18, mr: 18 }, base: { font: "Calibri", size: 10.5, line: 1.15, after: 4 },
          blocks: [
            p("title", "Sam Haddad", { sa: 2 }),
            p("subtitle", "Project Coordinator  ·  sam.haddad@mail.example  ·  +962 7 0000 0000  ·  Amman", { sa: 10 }),
            p("h2", "Profile"),
            p("normal", "Organised, calm under pressure and good with people. Five years coordinating projects for small teams, from first plan to final report."),
            p("h2", "Experience"),
            p("normal", [r("Project Coordinator", { b: true }), r("  —  Bright Path Consulting")], { sa: 0 }),
            p("normal", [r("2022 – present", { i: true, c: "#5a6372" })], { sa: 2 }),
            li("Ran weekly planning for 4 teams and kept 30+ projects on schedule."),
            li("Built a shared status report that leadership reads every Monday."),
            p("normal", [r("Office Administrator", { b: true }), r("  —  Cedar Clinic")], { sa: 0, sb: 6 }),
            p("normal", [r("2019 – 2022", { i: true, c: "#5a6372" })], { sa: 2 }),
            li("Managed bookings for 6 doctors and reduced no-shows with reminder calls."),
            li("Kept supplies, invoices and records in order."),
            p("h2", "Education"),
            p("normal", [r("B.A. Business Administration", { b: true }), r("  —  University of Jordan, 2019")]),
            p("h2", "Skills"),
            t,
            p("normal")
          ]
        };
      }
    },
    {
      id: "report", name: "Report", desc: "Title page, contents, sections", make: function () {
        var t = M.table(4, 3, [["Item", "Status", "Owner"], ["Survey", "Done", "Team A"], ["Interviews", "In progress", "Team B"], ["Final report", "Planned", "Team A"]]);
        t.cw = [0.45, 0.3, 0.25];
        return {
          title: "Report", hf: { header: "{title}", footer: "Page {page} of {pages}", ha: "right", fa: "center", first: true },
          blocks: [
            p("normal", "", { sb: 120 }),
            p("title", "Project Report"),
            p("subtitle", "Findings and next steps"),
            p("normal", [r("Prepared by: ", { b: true }), r("Your name")], { sb: 24 }),
            p("normal", [r("Date: ", { b: true }), r(today())]),
            M.pagebreak(),
            M.toc(),
            M.pagebreak(),
            p("h1", "1. Introduction"),
            p("normal", "Explain why this report exists, who it is for and what question it answers. Keep it to a short paragraph or two.", { align: "justify" }),
            p("h2", "1.1 Scope"),
            li("What is included"), li("What is not included"), li("The time period covered"),
            p("h1", "2. Method"),
            p("normal", "Describe how the information was gathered: surveys, interviews, data sources, and any limits the reader should know about.", { align: "justify" }),
            p("h1", "3. Findings"),
            p("normal", "Summarise the main results. A table often says it best:"),
            t,
            p("h1", "4. Recommendations"),
            li("First recommendation, with the reason behind it.", "ol"), li("Second recommendation.", "ol"), li("Third recommendation.", "ol"),
            p("h1", "5. Conclusion"),
            p("normal", "Close with the one thing the reader should remember.", { align: "justify" })
          ]
        };
      }
    },
    {
      id: "essay", name: "Essay (APA style)", desc: "Times 12, double spaced, title page", make: function () {
        var c = { align: "center" };
        return {
          title: "Essay", page: { size: "letter", mt: 25.4, mb: 25.4, ml: 25.4, mr: 25.4 }, base: { font: "Times New Roman", size: 12, line: 2, after: 0 },
          hf: { header: "{page}", footer: "", ha: "right", fa: "center", first: false },
          blocks: [
            p("normal", "", {}), p("normal", "", {}), p("normal", "", {}),
            p("normal", [r("The Title of Your Paper in Title Case", { b: true })], c),
            p("normal", "", c),
            p("normal", "Your Name", c),
            p("normal", "Department, University", c),
            p("normal", "Course Number: Course Name", c),
            p("normal", "Instructor Name", c),
            p("normal", today(), c),
            M.pagebreak(),
            p("normal", [r("The Title of Your Paper in Title Case", { b: true })], c),
            p("normal", "Begin the introduction here. Each paragraph starts with a half-inch first-line indent and the whole paper is double spaced, with one-inch margins and the page number at the top right.", { fi: "first" }),
            p("normal", "State your argument clearly near the end of the introduction, then support it in the body paragraphs that follow.", { fi: "first" }),
            p("normal", [r("First Main Point", { b: true })], c),
            p("normal", "Develop your first point with evidence and cite your sources in the text, for example (Author, 2024).", { fi: "first" }),
            p("normal", [r("Conclusion", { b: true })], c),
            p("normal", "Sum up the argument and why it matters.", { fi: "first" }),
            M.pagebreak(),
            p("normal", [r("References", { b: true })], c),
            p("normal", [r("Author, A. A., & Author, B. B. (2024). Title of the article. "), r("Journal Name, 12", { i: true }), r("(3), 45–67.")], { fi: "hang" }),
            p("normal", [r("Author, C. C. (2023). "), r("Title of the book", { i: true }), r(". Publisher.")], { fi: "hang" })
          ]
        };
      }
    },
    {
      id: "arletter", name: "خطاب رسمي", desc: "Arabic formal letter, right to left", make: function () {
        var R = { dir: "rtl" };
        return {
          title: "خطاب", base: { font: "Arial", size: 13, line: 1.5, after: 6 },
          blocks: [
            p("normal", [r("مؤسسة الأفق للخدمات", { b: true, z: 16, c: "#1d3557" })], { dir: "rtl", sa: 0 }),
            p("normal", "عمّان – المملكة الأردنية الهاشمية", { dir: "rtl", sa: 18 }),
            p("normal", "التاريخ: " + new Date().toLocaleDateString("ar"), { dir: "rtl", align: "right" }),
            p("normal", "الرقم: ٢٠٢٦/١٤", { dir: "rtl", align: "right", sa: 18 }),
            p("normal", [r("حضرة السيد/ ...................................... المحترم", { b: true })], R),
            p("normal", [r("الموضوع: ", { b: true }), r("طلب عرض أسعار")], { dir: "rtl", sa: 12 }),
            p("normal", "تحية طيبة وبعد،", R),
            p("normal", "نرجو التكرّم بتزويدنا بعرض أسعار لتوريد مستلزمات مكتبية لفرعنا الجديد، على أن يشمل العرض الأسعار ومدة التوريد وشروط الدفع، وذلك في موعد أقصاه نهاية الشهر الجاري.", { dir: "rtl", align: "justify" }),
            p("normal", "شاكرين لكم حسن تعاونكم، ونتطلع إلى استمرار العمل المشترك بيننا.", { dir: "rtl", align: "justify" }),
            p("normal", "وتفضلوا بقبول فائق الاحترام والتقدير،", { dir: "rtl", sb: 12, sa: 30 }),
            p("normal", [r("الاسم: ", { b: true }), r("..............................")], { dir: "rtl", align: "right" }),
            p("normal", [r("التوقيع: ", { b: true }), r("..............................")], { dir: "rtl", align: "right" })
          ]
        };
      }
    }
  ];
  window.DT = { list: T, make: function (id) { var t = T.filter(function (x) { return x.id === id; })[0] || T[0]; return M.newDoc(t.make()); } };
})();
