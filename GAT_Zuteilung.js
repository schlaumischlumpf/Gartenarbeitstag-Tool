    (() => {
      const SLOTS = ["S1", "S2", "S3", "S4"];
      const GRADES = [7, 8, 9, 10];
      const TRACKS = ["A", "B", "N", "E"];
      const ALL_CLASSES = GRADES.flatMap((grade) => TRACKS.map((track) => `${grade}${track}`));

      let studentIdCounter = 1;
      let projectIdCounter = 1;

      const state = {
        students: [],
        projects: [],
        assignments: {},
        conflicts: [],
        activeStep: 1,
        activeBoard: "regular",
        agProjects: [],
        agLists: [],
        agExtra: {},
        slotPriorities: {},
        activeStep1Tab: "students",
        filters: {
          search: "",
          className: "all",
          grade: "all"
        },
        pdf: {
          showHeader: true,
          showDate: true,
          showPageNumbers: true,
          schoolHeader: "Schule - Garbenarbeitstag",
          projectOrder: "id"
        },
        modal: {
          mode: null,
          entries: [],
          pendingResult: null
        },
        moveDialog: {
          studentId: null
        },
        agEditDialog: {
          studentId: null
        }
      };

      const LS_KEY = "gat_zuteilung_state";

      const byId = (id) => document.getElementById(id);

      function newStudentId() {
        const id = `s${studentIdCounter}`;
        studentIdCounter += 1;
        return id;
      }

      function newProjectId() {
        const id = `p${projectIdCounter}`;
        projectIdCounter += 1;
        return id;
      }

      function safeInt(value) {
        const parsed = Number.parseInt(value, 10);
        if (!Number.isFinite(parsed) || parsed < 0) {
          return 0;
        }
        return parsed;
      }

      function deepClone(value) {
        return JSON.parse(JSON.stringify(value));
      }

      function escapeHtml(value) {
        return String(value ?? "")
          .replaceAll("&", "&amp;")
          .replaceAll("<", "&lt;")
          .replaceAll(">", "&gt;")
          .replaceAll('"', "&quot;")
          .replaceAll("'", "&#039;");
      }

      function showMessage(text, type = "info") {
        const box = byId("message-box");
        box.className = `message-box show message-${type}`;
        box.textContent = text;
      }

      function hideMessage() {
        const box = byId("message-box");
        box.className = "message-box";
        box.textContent = "";
      }

      function buildEmptyDemands() {
        const demands = {};
        for (const grade of GRADES) {
          demands[grade] = {};
          for (const slot of SLOTS) {
            demands[grade][slot] = 0;
          }
        }
        return demands;
      }

      function emptyAgSlots() {
        return { S1: false, S2: false, S3: false, S4: false };
      }

      function normalizeAgSlots(value) {
        const slots = emptyAgSlots();
        if (value && typeof value === "object") {
          for (const slot of SLOTS) {
            slots[slot] = Boolean(value[slot]);
          }
        }
        return slots;
      }

      function getSelectedAgSlots(student) {
        return SLOTS.filter((slot) => Boolean(student.agSlots?.[slot]));
      }

      function classSortValue(className) {
        const match = String(className || "").match(/^(\d+)([A-Z])$/i);
        if (!match) {
          return 9999;
        }
        const grade = Number.parseInt(match[1], 10);
        const track = match[2].toUpperCase();
        return grade * 10 + Math.max(0, TRACKS.indexOf(track));
      }

      function compareStudentsByClassLast(a, b) {
        const classDiff = classSortValue(a.className) - classSortValue(b.className);
        if (classDiff !== 0) {
          return classDiff;
        }
        const last = a.lastName.localeCompare(b.lastName, "de", { sensitivity: "base" });
        if (last !== 0) {
          return last;
        }
        return a.firstName.localeCompare(b.firstName, "de", { sensitivity: "base" });
      }

      function getProjectById(projectId) {
        return state.projects.find((project) => project.id === projectId) || null;
      }

      function getStudentById(studentId) {
        return state.students.find((student) => student.id === studentId) || null;
      }

      function nextProjectNumber() {
        if (!state.projects.length) {
          return 1;
        }
        return Math.max(...state.projects.map((project) => safeInt(project.number))) + 1;
      }

      function createProject(number = nextProjectNumber(), name = "Neues Projekt") {
        return {
          id: newProjectId(),
          number,
          name,
          isSpecial: false,
          demands: buildEmptyDemands(),
          preference: {
            mode: "all",
            classes: {}
          }
        };
      }

      function normalizeProject(rawProject) {
        const project = createProject(safeInt(rawProject?.number) || nextProjectNumber(), rawProject?.name || "Projekt");
        if (rawProject?.id && typeof rawProject.id === "string") {
          project.id = rawProject.id;
        }
        project.name = String(rawProject?.name || "Projekt").trim() || "Projekt";
        project.number = safeInt(rawProject?.number) || nextProjectNumber();
        const sourceDemands = rawProject?.demands || {};
        for (const grade of GRADES) {
          for (const slot of SLOTS) {
            project.demands[grade][slot] = safeInt(sourceDemands?.[grade]?.[slot]);
          }
        }
        project.preference.mode = rawProject?.preference?.mode === "specific" ? "specific" : "all";
        project.preference.classes = {};
        if (rawProject?.preference?.classes && typeof rawProject.preference.classes === "object") {
          for (const [className, weight] of Object.entries(rawProject.preference.classes)) {
            if (ALL_CLASSES.includes(className) && safeInt(weight) > 0) {
              project.preference.classes[className] = safeInt(weight);
            }
          }
        }
        project.isSpecial = Boolean(rawProject?.isSpecial);
        return project;
      }

      function normalizeStudent(rawStudent) {
        const grade = safeInt(rawStudent?.grade || rawStudent?.Klassenstufe);
        const className = String(rawStudent?.className || rawStudent?.Klasse || "").trim().toUpperCase();
        const student = {
          id: String(rawStudent?.id || newStudentId()),
          firstName: String(rawStudent?.firstName || rawStudent?.Vorname || "").trim(),
          lastName: String(rawStudent?.lastName || rawStudent?.Nachname || "").trim(),
          className,
          grade,
          absent: Boolean(rawStudent?.absent),
          agMember: Boolean(rawStudent?.agMember),
          agProjectName: String(rawStudent?.agProjectName || "").trim(),
          agSlots: normalizeAgSlots(rawStudent?.agSlots)
        };
        if (student.agMember && !getSelectedAgSlots(student).length) {
          student.agSlots.S1 = true;
        }
        return student;
      }

      function recalcIdCounters() {
        let highestStudent = 0;
        let highestProject = 0;

        for (const student of state.students) {
          const match = student.id.match(/^(?:s)?(\d+)$/i);
          if (match) {
            highestStudent = Math.max(highestStudent, Number.parseInt(match[1], 10));
          }
        }

        for (const project of state.projects) {
          const match = project.id.match(/^(?:p)?(\d+)$/i);
          if (match) {
            highestProject = Math.max(highestProject, Number.parseInt(match[1], 10));
          }
        }

        studentIdCounter = highestStudent + 1;
        projectIdCounter = highestProject + 1;
      }

      function getAssignableStudents() {
        return state.students.filter((student) => !student.absent && !student.agMember);
      }

      function sanitizeAssignments() {
        const assignableIds = new Set(getAssignableStudents().map((student) => student.id));
        const projectIds = new Set(state.projects.map((project) => project.id));

        for (const [studentId, assignment] of Object.entries(state.assignments)) {
          if (!assignableIds.has(studentId) && !isSonderStudent(getStudentById(studentId))) {
            delete state.assignments[studentId];
            continue;
          }
          if (!assignment || !projectIds.has(assignment.projectId) || !SLOTS.includes(assignment.slot)) {
            delete state.assignments[studentId];
          }
        }
      }

      function getUnassignedStudents() {
        const assignedIds = new Set(Object.keys(state.assignments));
        return getAssignableStudents().filter((student) => !assignedIds.has(student.id));
      }

      function calculateDemandByGradeSlot() {
        const data = {};
        for (const grade of GRADES) {
          data[grade] = {};
          for (const slot of SLOTS) {
            data[grade][slot] = 0;
          }
        }

        for (const project of state.projects) {
          for (const grade of GRADES) {
            for (const slot of SLOTS) {
              data[grade][slot] += safeInt(project.demands?.[grade]?.[slot]);
            }
          }
        }

        return data;
      }

      function calculateAvailableByGradeClass() {
        const available = {};
        for (const grade of GRADES) {
          available[grade] = {};
          for (const track of TRACKS) {
            available[grade][`${grade}${track}`] = 0;
          }
        }

        for (const student of getAssignableStudents()) {
          if (!available[student.grade]) {
            continue;
          }
          if (!available[student.grade][student.className]) {
            available[student.grade][student.className] = 0;
          }
          available[student.grade][student.className] += 1;
        }

        return available;
      }

      function calculateAvailableByGrade() {
        const byClass = calculateAvailableByGradeClass();
        const byGrade = {};
        for (const grade of GRADES) {
          byGrade[grade] = Object.values(byClass[grade]).reduce((sum, value) => sum + value, 0);
        }
        return { byClass, byGrade };
      }

      function buildProjectSlotMap() {
        const map = {};
        for (const project of state.projects) {
          map[project.id] = {};
          for (const slot of SLOTS) {
            map[project.id][slot] = [];
          }
        }

        for (const [studentId, assignment] of Object.entries(state.assignments)) {
          const project = getProjectById(assignment.projectId);
          const student = getStudentById(studentId);
          if (!project || !student || !map[project.id] || !map[project.id][assignment.slot]) {
            continue;
          }
          map[project.id][assignment.slot].push(student);
        }

        for (const project of state.projects) {
          for (const slot of SLOTS) {
            map[project.id][slot].sort(compareStudentsByClassLast);
          }
        }

        return map;
      }

      function parseCsvLine(line, delimiter) {
        const cells = [];
        let current = "";
        let inQuotes = false;

        for (let index = 0; index < line.length; index += 1) {
          const char = line[index];
          const next = line[index + 1];
          if (char === '"') {
            if (inQuotes && next === '"') {
              current += '"';
              index += 1;
            } else {
              inQuotes = !inQuotes;
            }
            continue;
          }

          if (char === delimiter && !inQuotes) {
            cells.push(current.trim());
            current = "";
            continue;
          }

          current += char;
        }

        cells.push(current.trim());
        return cells;
      }

      function importStudentsFromCsv(csvText) {
        const lines = csvText.replace(/\r/g, "").split("\n").filter((line) => line.trim().length > 0);
        if (!lines.length) {
          throw new Error("Die CSV-Datei ist leer.");
        }

        const firstLine = lines[0].replace(/^\uFEFF/, "");
        const delimiter = firstLine.split(";").length > firstLine.split(",").length ? ";" : ",";
        const headers = parseCsvLine(firstLine, delimiter).map((header) => header.toLowerCase().replace(/\s+/g, ""));

        const required = ["vorname", "nachname", "klasse", "klassenstufe"];
        const missing = required.filter((field) => !headers.includes(field));
        if (missing.length) {
          throw new Error(`Fehlende CSV-Spalten: ${missing.join(", ")}`);
        }

        const indexMap = {
          firstName: headers.indexOf("vorname"),
          lastName: headers.indexOf("nachname"),
          className: headers.indexOf("klasse"),
          grade: headers.indexOf("klassenstufe")
        };

        const imported = [];
        let skipped = 0;

        for (let i = 1; i < lines.length; i += 1) {
          const cells = parseCsvLine(lines[i], delimiter);
          const firstName = String(cells[indexMap.firstName] || "").trim();
          const lastName = String(cells[indexMap.lastName] || "").trim();
          const className = String(cells[indexMap.className] || "").trim().toUpperCase();
          const grade = safeInt(cells[indexMap.grade]);

          if (!firstName || !lastName || !className || !GRADES.includes(grade)) {
            skipped += 1;
            continue;
          }

          imported.push(normalizeStudent({
            id: newStudentId(),
            firstName,
            lastName,
            className,
            grade,
            absent: false,
            agMember: false,
            agProjectName: "",
            agSlots: emptyAgSlots()
          }));
        }

        if (!imported.length) {
          throw new Error("Keine gültigen Schüler in der CSV gefunden.");
        }

        state.students = imported;
        state.assignments = {};
        state.conflicts = [];

        if (skipped > 0) {
          showMessage(`${imported.length} Schüler importiert. ${skipped} Zeilen wurden wegen unvollständiger Daten übersprungen.`, "warn");
        } else {
          showMessage(`${imported.length} Schüler wurden erfolgreich importiert.`, "ok");
        }
      }

      function loadSampleProjects() {
        const samples = [];

        function ensureProject(number, name) {
          let project = samples.find((item) => item.number === number);
          if (!project) {
            project = createProject(number, name);
            project.name = name;
            samples.push(project);
          }
          return project;
        }

        function setDemand(number, name, grade, s1, s2, s3, s4) {
          const project = ensureProject(number, name);
          if (grade === "alle") {
            for (const g of GRADES) {
              project.demands[g].S1 = s1;
              project.demands[g].S2 = s2;
              project.demands[g].S3 = s3;
              project.demands[g].S4 = s4;
            }
            return;
          }
          project.demands[grade].S1 = s1;
          project.demands[grade].S2 = s2;
          project.demands[grade].S3 = s3;
          project.demands[grade].S4 = s4;
        }

        setDemand(1, "Bodenbefreier und Rasenluefter", 7, 3, 3, 3, 0);
        setDemand(1, "Bodenbefreier und Rasenluefter", 8, 4, 4, 4, 0);
        setDemand(1, "Bodenbefreier und Rasenluefter", 9, 3, 3, 3, 0);

        setDemand(2, "Werkzeughandel", 7, 2, 2, 2, 2);
        setDemand(2, "Werkzeughandel", 8, 3, 2, 2, 3);

        setDemand(4, "Boesewichtejagd Teams", 7, 4, 4, 4, 0);
        setDemand(4, "Boesewichtejagd Teams", 8, 2, 2, 2, 0);
        setDemand(4, "Boesewichtejagd Teams", 9, 2, 2, 2, 0);
        setDemand(4, "Boesewichtejagd Teams", 10, 2, 2, 2, 0);

        setDemand(6, "Kompostkontrolle", "alle", 2, 2, 2, 2);
        setDemand(16, "WF-Beete Ausbau - Schrebergarten", "alle", 5, 5, 5, 5);
        setDemand(17, "Sportplatz fit machen", 7, 16, 15, 0, 0);

        state.projects = samples.map((project) => normalizeProject(project));
        sanitizeAssignments();
        state.conflicts = [];
        showMessage("Beispielprojekte wurden geladen.", "ok");
      }

      function getFilteredStudents() {
        const query = state.filters.search.trim().toLowerCase();
        return state.students.filter((student) => {
          if (state.filters.className !== "all" && student.className !== state.filters.className) {
            return false;
          }
          if (state.filters.grade !== "all" && String(student.grade) !== state.filters.grade) {
            return false;
          }
          if (!query) {
            return true;
          }
          const haystack = `${student.firstName} ${student.lastName} ${student.className}`.toLowerCase();
          return haystack.includes(query);
        }).sort(compareStudentsByClassLast);
      }

      function updateTopStats() {
        byId("stat-total").textContent = String(state.students.length);
        byId("stat-absent").textContent = String(state.students.filter((student) => student.absent).length);
        byId("stat-ag").textContent = String(state.students.filter((student) => student.agMember).length);
      }

      function renderStepNavigation() {
        const buttons = document.querySelectorAll("[data-action='goto-step']");
        const stepDone = [
          state.students.length > 0,
          state.projects.length > 0,
          Object.keys(state.assignments).length > 0,
          Object.keys(state.assignments).length > 0,
          true
        ];
        buttons.forEach((button) => {
          const step = Number.parseInt(button.dataset.step, 10);
          const isActive = step === state.activeStep;
          button.classList.toggle("active", isActive);
          button.classList.toggle("done", !isActive && stepDone[step - 1]);
        });

        for (let step = 1; step <= 5; step += 1) {
          byId(`step-${step}`).classList.toggle("active", step === state.activeStep);
        }
      }

      function renderClassFilterOptions() {
        const classFilter = byId("class-filter");
        const existing = [...new Set(state.students.map((student) => student.className).filter(Boolean))]
          .sort((a, b) => classSortValue(a) - classSortValue(b));
        const options = ["<option value=\"all\">Alle</option>"];
        for (const className of existing) {
          options.push(`<option value=\"${escapeHtml(className)}\">${escapeHtml(className)}</option>`);
        }
        classFilter.innerHTML = options.join("");
        classFilter.value = existing.includes(state.filters.className) ? state.filters.className : "all";
        if (classFilter.value === "all") {
          state.filters.className = "all";
        }
      }

      function renderStep1() {
        renderClassFilterOptions();

        const filtered = getFilteredStudents();
        const body = byId("students-body");
        if (!filtered.length) {
          body.innerHTML = "<tr><td colspan=\"8\" class=\"tiny\">Keine Schüler vorhanden oder Filter ohne Treffer.</td></tr>";
        } else {
          body.innerHTML = filtered.map((student) => {
            const agSlots = getSelectedAgSlots(student);
            return `
              <tr>
                <td>${escapeHtml(student.firstName)}</td>
                <td>${escapeHtml(student.lastName)}</td>
                <td>${escapeHtml(student.className)}</td>
                <td>${escapeHtml(String(student.grade))}</td>
                <td>
                  <input
                    type="checkbox"
                    data-action="toggle-absent"
                    data-student-id="${escapeHtml(student.id)}"
                    ${student.absent ? "checked" : ""}
                  >
                </td>
                <td>
                  <input
                    type="checkbox"
                    data-action="toggle-ag"
                    data-student-id="${escapeHtml(student.id)}"
                    ${student.agMember ? "checked" : ""}
                  >
                </td>
                <td>
                  <input
                    type="text"
                    data-action="ag-project"
                    data-student-id="${escapeHtml(student.id)}"
                    value="${escapeHtml(student.agProjectName || "")}"
                    placeholder="AG-Projekt"
                    ${student.agMember ? "" : "disabled"}
                  >
                </td>
                <td>
                  <div class="chips-row">
                    ${SLOTS.map((slot) => `
                      <label class="chip">
                        ${slot}
                        <input
                          type="checkbox"
                          data-action="ag-slot"
                          data-student-id="${escapeHtml(student.id)}"
                          data-slot="${slot}"
                          ${agSlots.includes(slot) ? "checked" : ""}
                          ${student.agMember ? "" : "disabled"}
                        >
                      </label>
                    `).join("")}
                  </div>
                </td>
              </tr>
            `;
          }).join("");
        }

        const absentCount = state.students.filter((student) => student.absent).length;
        const agCount = state.students.filter((student) => student.agMember).length;
        byId("step1-counts").innerHTML = `
          <strong>Uebersicht:</strong>
          ${state.students.length} Schüler gesamt,
          ${absentCount} abwesend/eingeschränkt,
          ${agCount} AG-Mitglieder,
          ${Math.max(0, state.students.length - absentCount - agCount)} normal verfügbar.
        `;
      }

      function renderProjectCard(project) {
        const prefMode = project.preference.mode;
        return `
          <article class="project-card" data-project-id="${escapeHtml(project.id)}">
            <div class="project-head">
              <label class="inline-field">Projekt-Nr.
                <input
                  type="number"
                  min="1"
                  data-action="project-number"
                  data-project-id="${escapeHtml(project.id)}"
                  value="${escapeHtml(String(project.number))}"
                >
              </label>
              <label class="inline-field">Projektname
                <input
                  type="text"
                  data-action="project-name"
                  data-project-id="${escapeHtml(project.id)}"
                  value="${escapeHtml(project.name)}"
                  placeholder="Projektname"
                >
              </label>
              <div class="btn-row">
                <label class="checkbox-row" style="font-size:.82rem; gap:6px; white-space:nowrap;">
                  <input
                    type="checkbox"
                    data-action="toggle-special"
                    data-project-id="${escapeHtml(project.id)}"
                    ${project.isSpecial ? "checked" : ""}
                  >
                  Sonderprojekt
                </label>
                <button type="button" class="danger" data-action="delete-project" data-project-id="${escapeHtml(project.id)}">Projekt löschen</button>
              </div>
            </div>

            <div class="project-demands table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Klassenstufe</th>
                    ${SLOTS.map((slot) => `<th>${slot}</th>`).join("")}
                  </tr>
                </thead>
                <tbody>
                  ${GRADES.map((grade) => `
                    <tr>
                      <td>Stufe ${grade}</td>
                      ${SLOTS.map((slot) => `
                        <td>
                          <input
                            type="number"
                            min="0"
                            value="${escapeHtml(String(project.demands[grade][slot]))}"
                            data-action="project-demand"
                            data-project-id="${escapeHtml(project.id)}"
                            data-grade="${grade}"
                            data-slot="${slot}"
                          >
                        </td>
                      `).join("")}
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>

            <div class="pref-panel">
              <label class="inline-field">Klassen-Präferenz
                <select data-action="pref-mode" data-project-id="${escapeHtml(project.id)}">
                  <option value="all" ${prefMode === "all" ? "selected" : ""}>Gleichmäßig aus allen Klassen</option>
                  <option value="specific" ${prefMode === "specific" ? "selected" : ""}>Bestimmte Klassen bevorzugen</option>
                </select>
              </label>

              ${prefMode === "specific" ? `
                <div class="tiny">Nur ausgewählte Klassen werden berücksichtigt. Gewichtung in Prozent (relative Verteilung).</div>
                <div class="pref-grid">
                  ${ALL_CLASSES.map((className) => {
                    const checked = Object.prototype.hasOwnProperty.call(project.preference.classes, className);
                    const weight = checked ? safeInt(project.preference.classes[className]) : "";
                    return `
                      <div class="pref-item">
                        <label class="checkbox-row">
                          <input
                            type="checkbox"
                            data-action="pref-class-toggle"
                            data-project-id="${escapeHtml(project.id)}"
                            data-class-name="${className}"
                            ${checked ? "checked" : ""}
                          >
                          ${className}
                        </label>
                        <input
                          type="number"
                          min="1"
                          max="100"
                          data-action="pref-class-weight"
                          data-project-id="${escapeHtml(project.id)}"
                          data-class-name="${className}"
                          value="${escapeHtml(String(weight))}"
                          placeholder="%"
                          ${checked ? "" : "disabled"}
                        >
                      </div>
                    `;
                  }).join("")}
                </div>
              ` : ""}
            </div>
          </article>
        `;
      }

      function renderDemandSummary() {
        const demand = calculateDemandByGradeSlot();
        const available = calculateAvailableByGrade();

        const tableRows = GRADES.map((grade) => {
          const totalDemand = SLOTS.reduce((sum, slot) => sum + demand[grade][slot], 0);
          const totalAvailable = available.byGrade[grade];
          // Konflikt: Gesamtbedarf der Stufe übersteigt verfügbare Schüler.
          const gradeConflict = totalDemand > totalAvailable;
          const rowClass = gradeConflict ? "status-err" : totalDemand > 0 ? "status-ok" : "";
          const cells = SLOTS.map((slot) => {
            const ist = demand[grade][slot];
            return `<td class="${gradeConflict ? "status-err" : ""}">${ist}</td>`;
          }).join("");
          return `
            <tr>
              <td>Stufe ${grade}</td>
              ${cells}
              <td class="${rowClass}">${totalDemand} / ${totalAvailable}</td>
            </tr>
          `;
        }).join("");

        byId("demand-summary").innerHTML = `
          <h3>Soll/Ist Uebersicht pro Klassenstufe und Slot</h3>
          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Stufe</th>
                  ${SLOTS.map((slot) => `<th>${slot}</th>`).join("")}
                  <th>Bedarf / Verfügbar</th>
                </tr>
              </thead>
              <tbody>
                ${tableRows}
              </tbody>
            </table>
          </div>
          <div class="tiny" style="margin-top:8px;">Rot: Gesamtbedarf der Klassenstufe überschreitet verfügbare Schüler.</div>
        `;
      }

      function renderStep2() {
        const container = byId("projects-container");
        if (!state.projects.length) {
          container.innerHTML = "<div class='summary-card tiny'>Noch keine Projekte vorhanden. Bitte Projekt hinzufügen oder Beispieldaten laden.</div>";
        } else {
          const sorted = [...state.projects].sort((a, b) => a.number - b.number || a.name.localeCompare(b.name, "de", { sensitivity: "base" }));
          container.innerHTML = sorted.map(renderProjectCard).join("");
        }
        renderDemandSummary();
      }

      function computePreConflicts() {
        const preConflicts = [];
        const demand = calculateDemandByGradeSlot();
        const available = calculateAvailableByGrade().byGrade;

        for (const grade of GRADES) {
          const gradeTotalDemand = SLOTS.reduce((sum, slot) => sum + demand[grade][slot], 0);
          if (gradeTotalDemand > available[grade]) {
            preConflicts.push(
              `Klassenstufe ${grade}: Gesamtbedarf ${gradeTotalDemand}, verfügbar nur ${available[grade]} Schüler.`
            );
          }
          // Kein per-Slot-Check hier: Einzelslot-Bedarf < Gesamtverfügbar wäre immer false
          // und zugleich implizit im Gesamtbedarf-Check enthalten.
        }

        return preConflicts;
      }

      function renderStep3Overview() {
        const available = calculateAvailableByGrade();
        const demand = calculateDemandByGradeSlot();

        const rows = GRADES.map((grade) => {
          const demandTotal = SLOTS.reduce((sum, slot) => sum + demand[grade][slot], 0);
          const availableTotal = available.byGrade[grade];

          let status = "ok";
          let statusText = "Gruen";
          if (demandTotal > availableTotal) {
            status = "err";
            statusText = "Rot";
          } else if (demandTotal > Math.floor(availableTotal * 0.9)) {
            status = "warn";
            statusText = "Gelb";
          }

          const classBreakdown = TRACKS.map((track) => `${grade}${track}: ${available.byClass[grade][`${grade}${track}`] || 0}`).join(" | ");

          return `
            <tr>
              <td>Stufe ${grade}</td>
              <td>${classBreakdown}</td>
              <td>${availableTotal}</td>
              <td>${demandTotal}</td>
              <td><span class="status-pill ${status}">${statusText}</span></td>
            </tr>
          `;
        }).join("");

        byId("overview-step3").innerHTML = `
          <h3>Verfügbarkeit und Bedarf</h3>
          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Stufe</th>
                  <th>Verfügbar je Klasse</th>
                  <th>Verfügbar gesamt</th>
                  <th>Bedarf gesamt</th>
                  <th>Ampel</th>
                </tr>
              </thead>
              <tbody>${rows}</tbody>
            </table>
          </div>
        `;
      }

      function renderStep3Conflicts() {
        const container = byId("conflicts-step3");
        if (!state.conflicts.length) {
          container.innerHTML = "<h3>Algorithmus-Konflikte</h3><div class='tiny'>Keine Konflikte aus der letzten Generierung.</div>";
          return;
        }

        container.innerHTML = `
          <h3>Algorithmus-Konflikte (${state.conflicts.length})</h3>
          <ul class="list-plain">
            ${state.conflicts.map((entry) => `<li>${escapeHtml(entry)}</li>`).join("")}
          </ul>
        `;
      }

      function renderPreviewTable() {
        const rows = Object.entries(state.assignments)
          .map(([studentId, assignment]) => {
            const student = getStudentById(studentId);
            const project = getProjectById(assignment.projectId);
            if (!student || !project) {
              return null;
            }
            return {
              student,
              project,
              slot: assignment.slot
            };
          })
          .filter(Boolean)
          .sort((a, b) => {
            const slotDiff = SLOTS.indexOf(a.slot) - SLOTS.indexOf(b.slot);
            if (slotDiff !== 0) {
              return slotDiff;
            }
            const projectDiff = a.project.name.localeCompare(b.project.name, "de", { sensitivity: "base" });
            if (projectDiff !== 0) {
              return projectDiff;
            }
            return compareStudentsByClassLast(a.student, b.student);
          });

        const body = byId("preview-body");
        if (!rows.length) {
          body.innerHTML = "<tr><td colspan=\"4\" class=\"tiny\">Noch keine Zuteilungen vorhanden.</td></tr>";
          return;
        }

        body.innerHTML = rows.map((row) => `
          <tr>
            <td>${escapeHtml(row.student.firstName)} ${escapeHtml(row.student.lastName)}</td>
            <td>${escapeHtml(row.student.className)}</td>
            <td>${escapeHtml(row.project.name)}</td>
            <td>${escapeHtml(row.slot)}</td>
          </tr>
        `).join("");
      }

      function renderStep3Priorities() {
        const container = byId("slot-priorities-step3");
        if (!container) {
          return;
        }

        const classesByGrade = {};
        for (const grade of GRADES) {
          classesByGrade[grade] = [...new Set(
            state.students
              .filter((s) => s.grade === grade)
              .map((s) => s.className)
              .filter(Boolean)
          )].sort();
        }

        const anyClasses = GRADES.some((g) => classesByGrade[g].length > 0);
        if (!anyClasses) {
          container.innerHTML = `
            <h3>Slot-Klassenprioritäten</h3>
            <p class="tiny">Keine Schüler importiert – bitte zuerst Schüler in Schritt 1 importieren.</p>
          `;
          return;
        }

        const slotHeaders = SLOTS.map((s) => `<th>${s}</th>`).join("");
        const gridRows = GRADES.map((grade) => {
          const cells = SLOTS.map((slot) => {
            const current = (state.slotPriorities[grade] || {})[slot] || "";
            const classes = classesByGrade[grade];
            const options = [
              `<option value="">— Keine —</option>`,
              ...classes.map((cn) => `<option value="${escapeHtml(cn)}"${current === cn ? " selected" : ""}>${escapeHtml(cn)}</option>`)
            ].join("");
            return `<td><select class="priority-select" data-priority-grade="${grade}" data-priority-slot="${escapeHtml(slot)}">${options}</select></td>`;
          }).join("");
          return `<tr><td class="priority-grade-label"><strong>Stufe ${grade}</strong></td>${cells}</tr>`;
        }).join("");

        container.innerHTML = `
          <h3>Slot-Klassenprioritäten</h3>
          <p class="tiny">Wähle pro Klassenstufe und Slot eine Klasse, die zuerst zugeteilt wird – bevor der Rest per Round-Robin verteilt wird. Hilfreich, wenn z.B. 9N in Slot 4 bevorzugt eingeteilt werden soll.</p>
          <div class="table-wrap" style="margin-top: 12px;">
            <table>
              <thead><tr><th>Stufe</th>${slotHeaders}</tr></thead>
              <tbody>${gridRows}</tbody>
            </table>
          </div>
        `;
      }

      function renderStep3() {
        renderStep3Overview();
        renderStep3Conflicts();
        renderStep3Priorities();
        renderPreviewTable();
      }

      function renderStudentChip(student) {
        return `
          <button
            type="button"
            class="student-chip"
            draggable="true"
            data-action="open-move"
            data-student-id="${escapeHtml(student.id)}"
            title="Klicken zum Umtragen, ziehen für Drag-and-Drop"
          >
            <span>${escapeHtml(student.firstName)} ${escapeHtml(student.lastName)}</span>
            <small>${escapeHtml(student.className)}</small>
          </button>
        `;
      }

      function renderStep4ClassFilter() {
        const sel = byId("step4-class-filter");
        if (!sel) return;
        const current = sel.value;
        const classes = [...new Set(state.students.map((s) => s.className).filter(Boolean))]
          .sort((a, b) => classSortValue(a) - classSortValue(b));
        sel.innerHTML = [
          "<option value=\"all\">Alle Klassen</option>",
          ...classes.map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`),
        ].join("");
        sel.value = classes.includes(current) ? current : "all";
      }

      function applyStep4Filter() {
        const search = (byId("step4-search")?.value ?? "").toLowerCase().trim();
        const grade  = byId("step4-grade-filter")?.value ?? "all";
        const cls    = byId("step4-class-filter")?.value ?? "all";
        const active = !!(search || grade !== "all" || cls !== "all");

        const activeBoard = document.querySelector(".board-section.active");
        if (!activeBoard) return;

        activeBoard.querySelectorAll(".student-chip").forEach((chip) => {
          if (!active) {
            chip.classList.remove("chip-highlight", "chip-dimmed");
            return;
          }
          const student = getStudentById(chip.dataset.studentId);
          if (!student) return;
          const fullName = `${student.firstName} ${student.lastName}`.toLowerCase();
          const matchName  = !search || fullName.includes(search) || student.className.toLowerCase().includes(search);
          const matchGrade = grade === "all" || String(student.grade) === grade;
          const matchClass = cls   === "all" || student.className === cls;
          if (matchName && matchGrade && matchClass) {
            chip.classList.add("chip-highlight");
            chip.classList.remove("chip-dimmed");
          } else {
            chip.classList.add("chip-dimmed");
            chip.classList.remove("chip-highlight");
          }
        });
      }

      /** Rendert eine Projektkarte mit 4 Slot-Spalten (shared by regular + sonder board). */
      function renderBoardProjectCard(project, slotMap) {
        const slotColumns = SLOTS.map((slot) => {
          const students = slotMap[project.id][slot] || [];
          const capacity = GRADES.reduce((sum, g) => sum + safeInt(project.demands?.[g]?.[slot]), 0);
          const fillClass = capacity === 0
            ? ""
            : students.length < capacity
              ? "status-warn"
              : students.length === capacity
                ? "status-ok"
                : "status-err";
          return `
            <section class="slot-col">
              <header>
                <span>${slot}</span>
                <span class="${fillClass}">${students.length} / ${capacity}</span>
              </header>
              <div class="dropzone" data-dropzone="true" data-target-project-id="${escapeHtml(project.id)}" data-target-slot="${slot}">
                ${students.length ? students.map(renderStudentChip).join("") : "<div class='tiny'>Leer</div>"}
              </div>
            </section>
          `;
        }).join("");
        const specialBadge = project.isSpecial ? " <span class='special-project-badge'>Sonder</span>" : "";
        return `
          <article class="slot-project${project.isSpecial ? " is-special" : ""}">
            <h3>${escapeHtml(String(project.number))} - ${escapeHtml(project.name)}${specialBadge}</h3>
            <div class="slot-grid">${slotColumns}</div>
          </article>
        `;
      }

      function renderStep4Board() {
        const board = byId("manual-board");
        const slotMap = buildProjectSlotMap();

        const regularProjects = [...state.projects]
          .filter((p) => !p.isSpecial)
          .sort((a, b) => a.number - b.number || a.name.localeCompare(b.name, "de", { sensitivity: "base" }));

        if (!regularProjects.length) {
          board.innerHTML = "<div class='summary-card tiny'>Es sind keine regulären Projekte vorhanden.</div>";
          byId("unassigned-zone").innerHTML = "";
        } else {
          board.innerHTML = regularProjects.map((p) => renderBoardProjectCard(p, slotMap)).join("");
        }

        const unassigned = getUnassignedStudents()
          .filter((s) => !s.agMember && !isSonderStudent(s))
          .sort(compareStudentsByClassLast);
        const countEl = byId("unassigned-count");
        countEl.textContent = String(unassigned.length);
        countEl.style.display = unassigned.length > 0 ? "inline-block" : "none";
        byId("unassigned-zone").innerHTML = unassigned.length
          ? unassigned.map(renderStudentChip).join("")
          : "<div class='tiny'>Keine übrigen Schüler.</div>";

        renderStep4ClassFilter();
        applyStep4Filter();
      }

      function renderSonderBoard() {
        const board = byId("sonder-board");
        if (!board) return;
        const slotMap = buildProjectSlotMap();

        const specialProjects = [...state.projects]
          .filter((p) => p.isSpecial)
          .sort((a, b) => a.number - b.number || a.name.localeCompare(b.name, "de", { sensitivity: "base" }));

        const sonderLists = state.agLists.filter((l) => l.isSpecial);

        const projectHtml = specialProjects.length
          ? specialProjects.map((p) => renderBoardProjectCard(p, slotMap)).join("")
          : "";

        const listHtml = sonderLists.map((list) => {
          const members = state.students
            .filter((s) => s.agMember && !s.absent && s.agProjectName === list.name)
            .sort(compareStudentsByClassLast);
          const slotCols = SLOTS.map((slot) => {
            const inSlot = members.filter((s) => s.agSlots?.[slot]);
            return `
              <section class="ag-slot-col">
                <header>
                  <span>${slot}</span>
                  <span class="ag-slot-count">${inSlot.length}</span>
                </header>
                <div class="ag-slot-body" data-dropzone="true" data-sonder-list-id="${escapeHtml(list.id)}" data-target-slot="${slot}">
                  ${inSlot.length
                    ? inSlot.map((s) => renderAgChip(s, false)).join("")
                    : "<div class='tiny' style='opacity:.5;'>–</div>"}
                </div>
              </section>
            `;
          }).join("");
          return `
            <article class="ag-project-card is-special">
              <div class="ag-card-head">
                <span class="ag-card-title">${escapeHtml(list.name)}</span>
                <span class="aglist-type-badge sonder">Sonder</span>
                <span class="tiny" style="color:var(--text-muted);">${members.length} Schüler</span>
              </div>
              <div class="ag-slot-grid">${slotCols}</div>
            </article>
          `;
        }).join("");

        if (!projectHtml && !listHtml) {
          board.innerHTML = "<div class='summary-card tiny'>Keine Sonderprojekte oder Sonderlisten vorhanden.</div>";
        } else {
          board.innerHTML = projectHtml + listHtml;
        }

        const sonderStudents = state.students
          .filter((s) => isSonderStudent(s) && !s.absent && !state.assignments[s.id])
          .sort(compareStudentsByClassLast);
        const countEl = byId("unassigned-sonder-count");
        if (countEl) {
          countEl.textContent = String(sonderStudents.length);
          countEl.style.display = sonderStudents.length > 0 ? "inline-block" : "none";
        }
        const zone = byId("unassigned-sonder-zone");
        if (zone) {
          zone.innerHTML = sonderStudents.length
            ? sonderStudents.map((s) => `<div class='sonder-chip'>${renderStudentChip(s)}</div>`).join("")
            : "<div class='tiny'>Keine ausstehenden Sonder-Schüler.</div>";
        }

        applyStep4Filter();
      }

      function renderStep4() {
        renderStep4Board();
        if (state.activeBoard === "ag") {
          renderAgBoard();
        } else if (state.activeBoard === "sonder") {
          renderSonderBoard();
        }
      }

      function renderStep5() {
        byId("opt-header").checked = state.pdf.showHeader;
        byId("opt-date").checked = state.pdf.showDate;
        byId("opt-pages").checked = state.pdf.showPageNumbers;
        byId("school-header").value = state.pdf.schoolHeader;
        byId("project-order").value = state.pdf.projectOrder;

        const assigned = Object.keys(state.assignments).length;
        const unassigned = getUnassignedStudents().length;
        const absent = state.students.filter((student) => student.absent).length;
        const ag = state.students.filter((student) => student.agMember).length;

        byId("export-summary").innerHTML = `
          <h3>Export-Status</h3>
          <div class="chips-row">
            <span class="chip">Regulär zugeteilt: ${assigned}</span>
            <span class="chip">Nicht zugeteilt: ${unassigned}</span>
            <span class="chip">Abwesend: ${absent}</span>
            <span class="chip">AG-Mitglieder: ${ag}</span>
          </div>
        `;
      }

      function renderAll() {
        sanitizeAssignments();
        updateTopStats();
        renderStepNavigation();
        renderStep1();
        renderStep1AgLists();
        renderStep2();
        renderStep3();
        renderStep4();
        renderStep5();
        saveState();
      }

      function openConflictModal(mode, title, subtitle, entries, pendingResult = null) {
        state.modal.mode = mode;
        state.modal.entries = entries;
        state.modal.pendingResult = pendingResult;

        byId("conflict-title").textContent = title;
        byId("conflict-subtitle").textContent = subtitle;
        byId("conflict-list").innerHTML = entries.map((entry) => `<li>${escapeHtml(entry)}</li>`).join("");
        byId("conflict-modal").classList.add("show");
      }

      function closeConflictModal() {
        byId("conflict-modal").classList.remove("show");
        state.modal.mode = null;
        state.modal.entries = [];
        state.modal.pendingResult = null;
      }

      function openMoveDialog(studentId) {
        const student = getStudentById(studentId);
        if (!student) {
          return;
        }
        state.moveDialog.studentId = studentId;

        const projectSelect = byId("move-project-select");
        const sortedProjects = [...state.projects].sort((a, b) => a.number - b.number || a.name.localeCompare(b.name, "de", { sensitivity: "base" }));
        projectSelect.innerHTML = [
          "<option value=''>Nicht zugeteilt</option>",
          ...sortedProjects.map((project) => `<option value='${escapeHtml(project.id)}'>${escapeHtml(String(project.number))} - ${escapeHtml(project.name)}</option>`)
        ].join("");

        const current = state.assignments[studentId] || null;
        projectSelect.value = current?.projectId || "";
        byId("move-slot-select").value = current?.slot || "S1";
        byId("move-slot-select").disabled = !current?.projectId;
        byId("move-student-label").textContent = `${student.firstName} ${student.lastName} (${student.className})`;

        byId("move-modal").classList.add("show");
      }

      function closeMoveDialog() {
        byId("move-modal").classList.remove("show");
        state.moveDialog.studentId = null;
      }

      // ── AG-Verwaltung ──────────────────────────────────────────────────

      /** Gibt alle bekannten AG-Projektnamen zurück (aus Schülerdaten + explizit angelegten).
       *  Sonderlisten (isSpecial) werden ausgeschlossen — sie erscheinen im Sonder-Tab. */
      function getAGProjectNames() {
        const sonderNames = new Set(
          state.agLists.filter((l) => l.isSpecial).map((l) => l.name)
        );
        const fromStudents = state.students
          .filter((student) => student.agMember && student.agProjectName.trim() && !sonderNames.has(student.agProjectName.trim()))
          .map((student) => student.agProjectName.trim());
        const all = [...new Set([
          ...state.agProjects.filter((name) => !sonderNames.has(name)),
          ...fromStudents
        ])];
        return all.sort((a, b) => a.localeCompare(b, "de", { sensitivity: "base" }));
      }

      /** Rendert einen einzelnen AG-Schüler-Chip. showSlots=false in den Slot-Spalten-Ansichten. */
      function renderAgChip(student, showSlots = true) {
        const slots = getSelectedAgSlots(student);
        const slotTags = showSlots
          ? slots.map((slot) => `<span class="ag-slot-tag">${slot}</span>`).join("")
          : "";
        return `
          <button
            type="button"
            class="student-chip ag-chip"
            draggable="true"
            data-action="open-ag-edit"
            data-student-id="${escapeHtml(student.id)}"
            title="Klicken zum Bearbeiten, ziehen für Drag-and-Drop"
          >
            <span>${escapeHtml(student.firstName)} ${escapeHtml(student.lastName)}</span>
            <span class="ag-chip-meta">
              <small>${escapeHtml(student.className)}</small>
              ${slotTags}
            </span>
          </button>
        `;
      }

      /** Rendert das komplette AG-Board mit Slot-Spalten je AG. */
      function renderAgBoard() {
        const board = byId("ag-board");
        if (!board) {
          return;
        }
        const agNames = getAGProjectNames();

        board.innerHTML = agNames.length
          ? agNames.map((name) => {
              const allStudents = state.students.filter(
                (s) => s.agMember && !s.absent && s.agProjectName.trim() === name
              ).sort(compareStudentsByClassLast);

              return `
                <article
                  class="ag-project-card"
                  data-dropzone="true"
                  data-ag-project="${escapeHtml(name)}"
                >
                  <div class="ag-card-head">
                    <span class="ag-card-title">${escapeHtml(name)}</span>
                    <span class="tiny" style="color:var(--text-muted);">${allStudents.length} Schüler</span>
                    <button
                      type="button"
                      class="danger"
                      style="padding:3px 8px; font-size:0.78rem;"
                      data-action="delete-ag-project"
                      data-ag-name="${escapeHtml(name)}"
                      title="AG-Projekt entfernen (Schüler bleiben AG-Mitglieder)"
                    >&#x2715;</button>
                  </div>
                  <div class="ag-card-body">
                    ${allStudents.length
                      ? allStudents.map((s) => renderAgChip(s)).join("")
                      : "<div class='tiny' style='opacity:.5;'>–</div>"}
                  </div>
                </article>
              `;
            }).join("")
          : "<div class='summary-card tiny'>Noch keine AGs vorhanden. AG-Schüler anlegen (Schritt 1) oder AG hinzufügen.</div>";

        // Unzugeordnete AG-Schüler
        const unassigned = state.students.filter(
          (student) => student.agMember && !student.absent && !student.agProjectName.trim()
        ).sort(compareStudentsByClassLast);
        const countEl = byId("unassigned-ag-count");
        if (countEl) {
          countEl.textContent = String(unassigned.length);
          countEl.style.display = unassigned.length > 0 ? "inline-block" : "none";
        }
        const unzoneEl = byId("unassigned-ag-zone");
        if (unzoneEl) {
          unzoneEl.innerHTML = unassigned.length
            ? unassigned.map(renderAgChip).join("")
            : "<div class='tiny'>Alle AG-Schüler sind einem Projekt zugewiesen.</div>";
        }
      }

      /** Verschiebt einen Schüler per Drag in ein anderes AG-Projekt. */
      function moveStudentToAg(studentId, targetAgName) {
        const student = getStudentById(studentId);
        if (!student || !student.agMember) {
          showMessage("Nur AG-Mitglieder können in die AG-Verwaltung verschoben werden.", "warn");
          return;
        }
        student.agProjectName = targetAgName;
        if (!getSelectedAgSlots(student).length) {
          student.agSlots.S1 = true;
        }
        renderAgBoard();
        renderStep1();
        renderStep5();
      }

      /** Verschiebt einen Sonder-Schüler in einen anderen Slot seiner Sonderliste. */
      function moveSonderStudentToSlot(studentId, listId, targetSlot) {
        const student = getStudentById(studentId);
        const list = state.agLists.find((l) => l.id === listId);
        if (!student || !list || !SLOTS.includes(targetSlot)) return;
        if (!student.agMember || student.agProjectName !== list.name) {
          showMessage("Schüler gehört nicht zu dieser Sonderliste.", "warn");
          return;
        }
        SLOTS.forEach((s) => { student.agSlots[s] = false; });
        student.agSlots[targetSlot] = true;
        renderSonderBoard();
        renderStep1AgLists();
        saveState();
      }

      /** Öffnet den AG-Bearbeiten-Dialog für einen Schüler. */
      function openAgEditDialog(studentId) {
        const student = getStudentById(studentId);
        if (!student || !student.agMember) {
          return;
        }
        state.agEditDialog.studentId = studentId;

        byId("ag-edit-student-label").textContent = `${student.firstName} ${student.lastName} (${student.className})`;

        const projectInput = byId("ag-edit-project-input");
        projectInput.value = student.agProjectName || "";

        // Datalist mit allen bekannten AG-Namen befüllen
        const dl = byId("ag-edit-projects-list");
        dl.innerHTML = getAGProjectNames()
          .map((name) => `<option value="${escapeHtml(name)}">`)
          .join("");

        // Slot-Checkboxen setzen
        byId("ag-edit-s1").checked = Boolean(student.agSlots?.S1);
        byId("ag-edit-s2").checked = Boolean(student.agSlots?.S2);
        byId("ag-edit-s3").checked = Boolean(student.agSlots?.S3);
        byId("ag-edit-s4").checked = Boolean(student.agSlots?.S4);

        byId("ag-edit-modal").classList.add("show");
      }

      /** Schließt den AG-Bearbeiten-Dialog. */
      function closeAgEditDialog() {
        byId("ag-edit-modal").classList.remove("show");
        state.agEditDialog.studentId = null;
      }

      /** Speichert die Änderungen aus dem AG-Bearbeiten-Dialog. */
      function confirmAgEdit() {
        const student = getStudentById(state.agEditDialog.studentId);
        if (!student) {
          closeAgEditDialog();
          return;
        }

        const newProjectName = byId("ag-edit-project-input").value.trim();
        student.agProjectName = newProjectName;

        student.agSlots.S1 = byId("ag-edit-s1").checked;
        student.agSlots.S2 = byId("ag-edit-s2").checked;
        student.agSlots.S3 = byId("ag-edit-s3").checked;
        student.agSlots.S4 = byId("ag-edit-s4").checked;

        // Mindestens ein Slot muss aktiv bleiben
        if (!getSelectedAgSlots(student).length) {
          student.agSlots.S1 = true;
          showMessage("Mindestens ein Slot muss aktiv bleiben — S1 wurde beibehalten.", "warn");
        }

        // Neues AG Projekt in state.agProjects merken (damit es auch nach letztem Schüler sichtbar bleibt)
        if (newProjectName && !state.agProjects.includes(newProjectName)) {
          state.agProjects.push(newProjectName);
        }

        closeAgEditDialog();
        renderAgBoard();
        renderStep1();
        renderStep5();
      }

      /** Board-Tab umschalten (Reguläre Zuteilung / AG-Verwaltung / Sonstiges). */
      function switchBoard(boardId) {
        state.activeBoard = boardId;
        document.querySelectorAll(".board-tab").forEach((tab) => {
          tab.classList.toggle("active", tab.dataset.board === boardId);
        });
        byId("board-regular").classList.toggle("active", boardId === "regular");
        byId("board-ag").classList.toggle("active", boardId === "ag");
        byId("board-sonder").classList.toggle("active", boardId === "sonder");
        if (boardId === "ag") {
          renderAgBoard();
        } else if (boardId === "sonder") {
          renderSonderBoard();
        }
        applyStep4Filter();
      }

      // ── AG-Listen Hilfsfunktionen ──────────────────────────────────
      let agListIdCounter = 1;

      function createAgList(name, isSpecial = false) {
        const id = `al${agListIdCounter}`;
        agListIdCounter += 1;
        return { id, name, isSpecial };
      }

      function isSonderStudent(student) {
        if (!student?.agMember) {
          return false;
        }
        const list = state.agLists.find((l) => l.name === student.agProjectName);
        return list ? list.isSpecial : false;
      }

      function addStudentToAgList(studentId, listId) {
        const student = getStudentById(studentId);
        const list = state.agLists.find((l) => l.id === listId);
        if (!student || !list) {
          return;
        }
        student.agMember = true;
        student.agProjectName = list.name;
        if (!getSelectedAgSlots(student).length) {
          student.agSlots.S1 = true;
        }
        delete state.assignments[studentId];
        renderAll();
      }

      function removeStudentFromAgList(studentId) {
        const student = getStudentById(studentId);
        if (!student) {
          return;
        }
        student.agMember = false;
        student.agProjectName = "";
        student.agSlots = emptyAgSlots();
        renderAll();
      }

      function addStudentToAgExtra(studentId, listId) {
        if (!state.agExtra[listId]) state.agExtra[listId] = [];
        if (!state.agExtra[listId].includes(studentId)) {
          state.agExtra[listId].push(studentId);
        }
        renderStep1AgLists();
        saveState();
      }

      function removeStudentFromAgExtra(studentId, listId) {
        if (!state.agExtra[listId]) return;
        state.agExtra[listId] = state.agExtra[listId].filter((id) => id !== studentId);
        renderStep1AgLists();
        saveState();
      }

      function getExtraStudentsForList(listId) {
        return (state.agExtra[listId] || [])
          .map((id) => getStudentById(id))
          .filter((s) => s && !s.absent);
      }

      function renderStep1AgLists() {
        const board = byId("step1-aglist-board");
        if (!board) {
          return;
        }

        const poolStudents = state.students.filter((s) => !s.agMember && !s.absent)
          .sort(compareStudentsByClassLast);

        const poolChips = poolStudents.map((s) => `
          <span
            class="pool-chip"
            draggable="true"
            data-student-id="${escapeHtml(s.id)}"
            data-drag-source="aglist-pool"
          >${escapeHtml(s.firstName)} ${escapeHtml(s.lastName)} <small>${escapeHtml(s.className)}</small></span>
        `).join("");

        const agListCards = state.agLists.map((list) => {
          const members = state.students
            .filter((s) => s.agMember && s.agProjectName === list.name)
            .sort(compareStudentsByClassLast);
          const chips = members.map((s) => `
            <button
              type="button"
              class="student-chip ${list.isSpecial ? "sonder-chip-inner" : ""}"
              data-action="remove-from-aglist"
              data-student-id="${escapeHtml(s.id)}"
              title="Aus Liste entfernen"
            >${escapeHtml(s.firstName)} ${escapeHtml(s.lastName)} <small>${escapeHtml(s.className)}</small></button>
          `).join("");
          const extraMembers = getExtraStudentsForList(list.id).sort(compareStudentsByClassLast);
          const extraChips = extraMembers.map((s) => `
            <button
              type="button"
              class="student-chip extra-chip"
              data-action="remove-from-ag-extra"
              data-student-id="${escapeHtml(s.id)}"
              data-aglist-id="${escapeHtml(list.id)}"
              title="Aus dieser Liste entfernen (dupliziert)"
            >${escapeHtml(s.firstName)} ${escapeHtml(s.lastName)} <small>${escapeHtml(s.className)}</small></button>
          `).join("");
          const totalCount = members.length + extraMembers.length;
          return `
            <div class="aglist-card type-${list.isSpecial ? "sonder" : "ag"}">
              <div class="aglist-card-head">
                <span class="aglist-card-title">${escapeHtml(list.name)}</span>
                <span class="aglist-type-badge ${list.isSpecial ? "sonder" : "ag"}">${list.isSpecial ? "Sonder" : "AG"}</span>
                <span class="tiny">${totalCount} Schüler</span>
                <button
                  type="button"
                  class="danger"
                  style="padding:3px 8px; font-size:.76rem;"
                  data-action="delete-aglist"
                  data-aglist-id="${escapeHtml(list.id)}"
                  title="Liste löschen"
                >&#x2715;</button>
              </div>
              <div
                class="aglist-dropzone"
                data-dropzone="true"
                data-aglist-id="${escapeHtml(list.id)}"
              >${chips}${extraChips}${(!chips && !extraChips) ? "<div class='tiny'>Schüler hierher ziehen&hellip;</div>" : ""}</div>
            </div>
          `;
        }).join("");

        board.innerHTML = `
          <div class="aglist-layout">
            <div class="aglist-pool">
              <h3>Schüler-Pool (${poolStudents.length})</h3>
              <input
                id="aglist-pool-search"
                type="search"
                placeholder="Suche nach Name oder Klasse…"
                autocomplete="off"
              />
              <div
                class="aglist-pool-zone"
                data-dropzone="true"
                data-aglist-pool="true"
              >${poolChips || "<div class='tiny'>Alle Schüler sind zugeteilt.</div>"}</div>
            </div>
            <div class="aglist-groups">
              <div class="aglist-groups-header">
                <h3>AG-Gruppen &amp; Sonderlisten</h3>
                <div class="btn-row">
                  <button type="button" class="primary" data-action="add-ag-list">AG hinzufügen</button>
                  <button type="button" data-action="add-sonder-list">Sonderliste erstellen</button>
                </div>
              </div>
              ${agListCards || "<div class='tiny' style='padding:8px;'>Noch keine Listen. \"AG hinzufügen\" oder \"Sonderliste erstellen\" klicken.</div>"}
            </div>
          </div>
        `;
      }

      function moveStudent(studentId, targetProjectId, targetSlot) {
        const student = getStudentById(studentId);
        if (!student) {
          return;
        }
        if (student.absent || (student.agMember && !isSonderStudent(student))) {
          showMessage("Abwesende oder AG-Mitglieder sind für die normale Zuteilung gesperrt.", "warn");
          return;
        }

        if (!targetProjectId) {
          delete state.assignments[studentId];
        } else {
          const project = getProjectById(targetProjectId);
          if (!project || !SLOTS.includes(targetSlot)) {
            showMessage("Ungültiges Ziel für Umtragung.", "err");
            return;
          }
          state.assignments[studentId] = { projectId: targetProjectId, slot: targetSlot };
        }

        renderStep3();
        renderStep4();
        renderStep5();
      }

      function randomShuffle(array) {
        const result = [...array];
        for (let i = result.length - 1; i > 0; i -= 1) {
          const j = Math.floor(Math.random() * (i + 1));
          [result[i], result[j]] = [result[j], result[i]];
        }
        return result;
      }

      function runAllocation() {
        const assignable = randomShuffle(getAssignableStudents());

        const pools = {};
        for (const grade of GRADES) {
          pools[grade] = {};
          for (const track of TRACKS) {
            pools[grade][`${grade}${track}`] = [];
          }
        }

        for (const student of assignable) {
          if (!pools[student.grade]) {
            continue;
          }
          if (!pools[student.grade][student.className]) {
            pools[student.grade][student.className] = [];
          }
          pools[student.grade][student.className].push(student.id);
        }

        for (const grade of GRADES) {
          for (const className of Object.keys(pools[grade])) {
            pools[grade][className] = randomShuffle(pools[grade][className]);
          }
        }

        // ── Priority-Reservierung ────────────────────────────────────────────
        // Schüler der konfigurierten Prioritätsklassen werden VORAB aus dem
        // allgemeinen Pool herausgenommen. Round-Robin für andere Slots greift
        // dann nicht mehr auf sie zu, sodass sie garantiert in ihren
        // Prioritäts-Slots landen.
        const priorityReserve = {};     // grade -> className -> [studentIds]
        const priorityClassSlots = {};  // grade -> className -> Set<slot>

        for (const grade of GRADES) {
          priorityReserve[grade] = {};
          priorityClassSlots[grade] = {};
          const gradeConfig = state.slotPriorities[grade] || {};
          const classQuota = {};

          for (const [slot, className] of Object.entries(gradeConfig)) {
            if (!className || !pools[grade]?.[className]) continue;
            if (!priorityClassSlots[grade][className]) {
              priorityClassSlots[grade][className] = new Set();
            }
            priorityClassSlots[grade][className].add(slot);
            const demand = state.projects
              .filter((p) => !p.isSpecial)
              .reduce((sum, p) => sum + safeInt(p.demands?.[grade]?.[slot]), 0);
            classQuota[className] = (classQuota[className] || 0) + demand;
          }

          for (const [className, quota] of Object.entries(classQuota)) {
            const bucket = pools[grade][className] || [];
            const count = Math.min(quota, bucket.length);
            // splice vom Ende: die zuletzt gemischten Schüler reservieren
            priorityReserve[grade][className] = bucket.splice(bucket.length - count);
          }
        }
        // ────────────────────────────────────────────────────────────────────

        const rrPointer = {};
        for (const grade of GRADES) {
          rrPointer[grade] = Math.floor(Math.random() * TRACKS.length);
        }

        function totalAvailableInGrade(grade) {
          return Object.values(pools[grade]).reduce((sum, list) => sum + list.length, 0);
        }

        function pickRoundRobin(grade, need, slot = "") {
          // Klassen, die noch reservierte Schüler für ANDERE Slots haben, überspringen.
          // Nur wenn der aktuelle Slot einer ihrer Prioritäts-Slots ist, dürfen sie
          // via Round-Robin eingeplant werden (dann aber als Ergänzung zum Reserve-Pool).
          const skipClasses = new Set();
          for (const [className, reserved] of Object.entries(priorityReserve[grade] || {})) {
            if (reserved.length > 0 && !(priorityClassSlots[grade]?.[className]?.has(slot))) {
              skipClasses.add(className);
            }
          }

          const picked = [];
          const cycle = TRACKS
            .map((track) => `${grade}${track}`)
            .filter((cn) => !skipClasses.has(cn));

          if (!cycle.length) return picked;

          let guard = 0;
          while (picked.length < need && totalAvailableInGrade(grade) > 0 && guard < 4000) {
            const className = cycle[rrPointer[grade] % cycle.length];
            rrPointer[grade] += 1;
            const bucket = pools[grade][className] || [];
            if (bucket.length) {
              picked.push(bucket.pop());
            }
            guard += 1;
          }

          while (picked.length < need) {
            const fallback = cycle.find((className) => (pools[grade][className] || []).length > 0);
            if (!fallback) {
              break;
            }
            picked.push(pools[grade][fallback].pop());
          }

          return picked;
        }

        function pickWeightedClass(classEntries) {
          const totalWeight = classEntries.reduce((sum, entry) => sum + entry.weight, 0);
          if (totalWeight <= 0) {
            return classEntries[0]?.className || null;
          }
          let random = Math.random() * totalWeight;
          for (const entry of classEntries) {
            random -= entry.weight;
            if (random <= 0) {
              return entry.className;
            }
          }
          return classEntries[classEntries.length - 1]?.className || null;
        }

        function pickWeighted(grade, need, classWeights, slot = "") {
          const selected = Object.entries(classWeights)
            .filter(([className, weight]) => className.startsWith(String(grade)) && safeInt(weight) > 0)
            .map(([className, weight]) => ({ className, weight: safeInt(weight) }));

          if (!selected.length) {
            return [];
          }

          const totalWeight = selected.reduce((sum, entry) => sum + entry.weight, 0);
          if (totalWeight <= 0) {
            return [];
          }

          const quotas = {};
          let assignedQuota = 0;
          const remainders = [];

          for (const entry of selected) {
            const raw = (entry.weight / totalWeight) * need;
            const base = Math.floor(raw);
            quotas[entry.className] = base;
            assignedQuota += base;
            remainders.push({ className: entry.className, rest: raw - base });
          }

          let open = need - assignedQuota;
          remainders.sort((a, b) => b.rest - a.rest);
          // Largest-Remainder-Methode: die open Klassen mit den größten Resten bekommen +1.
          for (let i = 0; i < open && i < remainders.length; i += 1) {
            quotas[remainders[i].className] += 1;
          }

          const picked = [];
          for (const entry of selected) {
            const bucket = pools[grade][entry.className] || [];
            const take = Math.min(quotas[entry.className], bucket.length);
            for (let i = 0; i < take; i += 1) {
              picked.push(bucket.pop());
            }
          }

          // Wenn Quotierung nicht ausreicht, wird aus den selektierten Klassen gewichtet nachgezogen.
          let guard = 0;
          while (picked.length < need && guard < 4000) {
            const candidates = selected
              .filter((entry) => (pools[grade][entry.className] || []).length > 0)
              .map((entry) => ({ className: entry.className, weight: entry.weight }));
            if (!candidates.length) {
              break;
            }
            const className = pickWeightedClass(candidates);
            if (!className) {
              break;
            }
            picked.push(pools[grade][className].pop());
            guard += 1;
          }

          // Fallback: Wenn spezifizierte Klassen komplett erschöpft sind, aus den restlichen
          // Klassen der Stufe per Round-Robin nachziehen, um unnötige Konflikte zu vermeiden.
          if (picked.length < need) {
            const fallback = pickRoundRobin(grade, need - picked.length, slot);
            picked.push(...fallback);
          }

          return picked;
        }

        // Wenn für grade+slot eine Prioritätklasse konfiguriert ist, werden zuerst
        // die reservierten Schüler dieser Klasse verwendet, dann restliche aus dem
        // allgemeinen Pool, und schließlich Round-Robin für verbleibende Plätze.
        function pickWithPriority(grade, need, slot) {
          const priorityClass = (state.slotPriorities[grade] || {})[slot] || "";
          if (!priorityClass) {
            return pickRoundRobin(grade, need, slot);
          }

          const reserved = (priorityReserve[grade] || {})[priorityClass] || [];
          const mainBucket = (pools[grade] || {})[priorityClass] || [];
          const picked = [];

          // 1. Aus dem reservierten Pool (wurden vorab gesichert)
          while (picked.length < need && reserved.length > 0) {
            picked.push(reserved.pop());
          }
          // 2. Aus dem allgemeinen Pool (falls Reserve aufgebraucht)
          while (picked.length < need && mainBucket.length > 0) {
            picked.push(mainBucket.pop());
          }
          // 3. Rest per Round-Robin (Prioritätsklasse ist jetzt erschöpft → wird übersprungen)
          if (picked.length < need) {
            picked.push(...pickRoundRobin(grade, need - picked.length, slot));
          }
          return picked;
        }

        const resultAssignments = {};
        const conflictDetails = [];

        // Kernalgorithmus: Slot für Slot, Projekt für Projekt, Stufe für Stufe.
        // Projektreihenfolge wird pro Slot neu gemischt → kein Projekt wird systematisch bevorzugt.
        for (const slot of SLOTS) {
          // Round-Robin-Pointer pro Slot zurücksetzen: jeder Slot startet mit frischer Klassenverteilung.
          for (const grade of GRADES) {
            rrPointer[grade] = Math.floor(Math.random() * TRACKS.length);
          }
          const slotProjects = randomShuffle([...state.projects]);
          for (const project of slotProjects) {
            if (project.isSpecial) {
              continue; // Sonderprojekte werden nicht automatisch befüllt
            }
            for (const grade of GRADES) {
              const need = safeInt(project.demands?.[grade]?.[slot]);
              if (need <= 0) {
                continue;
              }

              let picked = [];
              if (project.preference.mode === "specific") {
                picked = pickWeighted(grade, need, project.preference.classes, slot);
              } else {
                picked = pickWithPriority(grade, need, slot);
              }

              if (picked.length < need) {
                conflictDetails.push(
                  `Projekt \"${project.name}\", Slot ${slot}, Klassenstufe ${grade}: Bedarf ${need}, verfügbar nur ${picked.length}.`
                );
              }

              for (const studentId of picked) {
                resultAssignments[studentId] = { projectId: project.id, slot };
              }
            }
          }
        }

        return {
          assignments: resultAssignments,
          conflicts: conflictDetails
        };
      }

      function executeGeneration() {
        const btn = byId("btn-generate");
        btn.disabled = true;
        btn.textContent = "Wird berechnet…";
        // Kurze Verzögerung damit der Browser den Disabled-State darstellen kann.
        setTimeout(() => {
          try {
            _runGeneration();
          } finally {
            btn.disabled = false;
            btn.textContent = "Zuteilung generieren";
          }
        }, 20);
      }

      function _runGeneration() {
        if (!state.students.length) {
          showMessage("Bitte zuerst Schüler importieren.", "warn");
          return;
        }
        if (!state.projects.length) {
          showMessage("Bitte zuerst Projekte anlegen.", "warn");
          return;
        }

        const preConflicts = computePreConflicts();
        if (preConflicts.length) {
          openConflictModal(
            "pre",
            "Unterbesetzungs-Warnung",
            "Mindestens eine Slot+Stufe-Kombination ist voraussichtlich unterbesetzt.",
            preConflicts
          );
          return;
        }

        finalizeGeneration();
      }

      function finalizeGeneration() {
        const result = runAllocation();
        if (result.conflicts.length) {
          openConflictModal(
            "post",
            "Konflikte nach Algorithmuslauf",
            "Nicht alle Bedarfe konnten vollständig gedeckt werden.",
            result.conflicts,
            result
          );
          return;
        }

        state.assignments = result.assignments;
        state.conflicts = result.conflicts;
        showMessage("Zuteilung wurde erzeugt.", "ok");
        renderAll();
      }

      function downloadJsonSnapshot() {
        const payload = {
          version: 2,
          exportedAt: new Date().toISOString(),
          students: state.students,
          projects: state.projects,
          assignments: state.assignments,
          agProjects: state.agProjects,
          agLists: state.agLists,
          agExtra: state.agExtra,
          slotPriorities: state.slotPriorities,
          pdf: state.pdf
        };
        const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = "GAT_Zuteilung.json";
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(url);
        showMessage("JSON-Datei wurde gespeichert.", "ok");
      }

      function loadJsonSnapshot(text) {
        let parsed;
        try {
          parsed = JSON.parse(text);
        } catch (error) {
          throw new Error("JSON-Datei ist ungültig.");
        }

        if (!Array.isArray(parsed.students) || !Array.isArray(parsed.projects)) {
          throw new Error("JSON-Struktur passt nicht (students/projects fehlen).");
        }

        state.students = parsed.students.map(normalizeStudent);
        state.projects = parsed.projects.map(normalizeProject);
        state.assignments = typeof parsed.assignments === "object" && parsed.assignments ? deepClone(parsed.assignments) : {};
        state.agProjects = Array.isArray(parsed.agProjects)
          ? parsed.agProjects.filter((name) => typeof name === "string" && name.trim())
          : [];
        state.agLists = Array.isArray(parsed.agLists)
          ? parsed.agLists.filter((l) => l && typeof l.id === "string" && typeof l.name === "string")
          : [];
        state.agExtra = (parsed.agExtra && typeof parsed.agExtra === "object" && !Array.isArray(parsed.agExtra))
          ? parsed.agExtra
          : {};
        state.pdf = {
          ...state.pdf,
          ...(parsed.pdf || {})
        };
        state.slotPriorities = (parsed.slotPriorities && typeof parsed.slotPriorities === "object" && !Array.isArray(parsed.slotPriorities))
          ? parsed.slotPriorities
          : {};
        state.conflicts = [];

        recalcIdCounters();
        sanitizeAssignments();
      }

      function ensureJsPdf() {
        if (!window.jspdf || !window.jspdf.jsPDF) {
          throw new Error("jsPDF wurde nicht geladen. Bitte Internetverbindung prüfen.");
        }
        return window.jspdf.jsPDF;
      }

      function applyPdfDecorations(doc, title) {
        const totalPages = doc.getNumberOfPages();
        const pageWidth = doc.internal.pageSize.getWidth();
        const pageHeight = doc.internal.pageSize.getHeight();
        const dateText = new Date().toLocaleDateString("de-DE");

        for (let page = 1; page <= totalPages; page += 1) {
          doc.setPage(page);

          if (state.pdf.showHeader) {
            doc.setFontSize(10);
            doc.setTextColor(40, 68, 55);
            doc.text(state.pdf.schoolHeader || "GAT", 14, 10);
            doc.text(title, pageWidth - 14, 10, { align: "right" });
            doc.setDrawColor(180, 201, 183);
            doc.line(14, 12, pageWidth - 14, 12);
          }

          if (state.pdf.showDate) {
            doc.setFontSize(9);
            doc.setTextColor(85, 100, 90);
            doc.text(dateText, 14, pageHeight - 8);
          }

          if (state.pdf.showPageNumbers) {
            doc.setFontSize(9);
            doc.setTextColor(85, 100, 90);
            doc.text(`Seite ${page} / ${totalPages}`, pageWidth - 14, pageHeight - 8, { align: "right" });
          }
        }
      }

      function collectProjectExportRecords() {
        const slotMap = buildProjectSlotMap();
        const regularRecords = state.projects.map((project) => ({
          key: `regular:${project.id}`,
          number: project.number,
          name: project.name,
          slots: deepClone(slotMap[project.id] || { S1: [], S2: [], S3: [], S4: [] })
        }));

        const regularByName = new Map();
        for (const record of regularRecords) {
          regularByName.set(record.name.trim().toLowerCase(), record);
        }

        const agOnly = new Map();
        for (const student of state.students) {
          if (!student.agMember || student.absent) {
            continue;
          }
          const projectName = student.agProjectName.trim() || "AG ohne Projekt";
          const slots = getSelectedAgSlots(student);
          const targetSlots = slots.length ? slots : ["S1"];
          const existingRegular = regularByName.get(projectName.toLowerCase());

          if (existingRegular) {
            for (const slot of targetSlots) {
              existingRegular.slots[slot].push(student);
            }
          } else {
            if (!agOnly.has(projectName)) {
              agOnly.set(projectName, {
                key: `ag:${projectName}`,
                number: 9999,
                name: projectName,
                slots: { S1: [], S2: [], S3: [], S4: [] }
              });
            }
            const record = agOnly.get(projectName);
            for (const slot of targetSlots) {
              record.slots[slot].push(student);
            }
          }
        }

        const allRecords = [...regularRecords, ...agOnly.values()];
        for (const record of allRecords) {
          for (const slot of SLOTS) {
            record.slots[slot].sort(compareStudentsByClassLast);
          }
        }

        if (state.pdf.projectOrder === "name") {
          allRecords.sort((a, b) => a.name.localeCompare(b.name, "de", { sensitivity: "base" }));
        } else {
          allRecords.sort((a, b) => a.number - b.number || a.name.localeCompare(b.name, "de", { sensitivity: "base" }));
        }

        return allRecords;
      }

      function exportProjectPdf() {
        const JsPdf = ensureJsPdf();
        const records = collectProjectExportRecords();
        if (!records.length) {
          showMessage("Keine Projekte für PDF vorhanden.", "warn");
          return;
        }

        const doc = new JsPdf({ unit: "mm", format: "a4" });
        const PH = doc.internal.pageSize.getHeight();
        const ML = 14; // linker Rand

        // Spaltenbreiten: ☐ | Nr. | Nachname | Vorname | Klasse
        const CW = [8, 11, 72, 60, 20];
        const TW = CW.reduce((s, w) => s + w, 0);
        const ROW_H = 7.2;

        records.forEach((record, recIdx) => {
          if (recIdx > 0) { doc.addPage(); }
          let y = 26;
          let rowNr = 0;

          const titelBand = (label) => {
            doc.setFillColor(219, 240, 224);
            doc.rect(ML, y - 6.5, TW, 10, "F");
            doc.setFontSize(13);
            doc.setFont("helvetica", "bold");
            doc.setTextColor(22, 52, 35);
            doc.text(label, ML + 3, y);
            doc.setFont("helvetica", "normal");
            y += 8;
          };

          const kopfzeile = () => {
            doc.setFillColor(195, 222, 200);
            doc.rect(ML, y - 5.5, TW, 8, "F");
            doc.setFontSize(8.5);
            doc.setFont("helvetica", "bold");
            doc.setTextColor(28, 56, 36);
            let x = ML;
            ["", "Nr.", "Nachname", "Vorname", "Klasse"].forEach((l, i) => {
              if (l) { doc.text(l, x + 2.5, y); }
              x += CW[i];
            });
            doc.setFont("helvetica", "normal");
            doc.setDrawColor(148, 188, 156);
            doc.line(ML, y + 2, ML + TW, y + 2);
            y += 8;
          };

          const slotBand = (slot, count) => {
            doc.setFillColor(236, 248, 238);
            doc.rect(ML, y - 5, TW, 8, "F");
            doc.setFontSize(9.5);
            doc.setFont("helvetica", "bold");
            doc.setTextColor(38, 88, 52);
            doc.text(`Slot ${slot}  \u2014  ${count} Sch\u00FCler`, ML + 4, y);
            doc.setFont("helvetica", "normal");
            y += 8;
          };

          const schuelerZeile = (student) => {
            rowNr++;
            if (rowNr % 2 === 0) {
              doc.setFillColor(250, 254, 251);
              doc.rect(ML, y - 4.8, TW, ROW_H, "F");
            }
            let x = ML;
            doc.setDrawColor(52, 68, 56);
            doc.setFillColor(255, 255, 255);
            doc.rect(x + 1.8, y - 3.2, 4, 4, "FD");
            x += CW[0];
            doc.setFontSize(9);
            doc.setTextColor(36, 48, 40);
            doc.text(String(rowNr), x + 2.5, y);
            x += CW[1];
            doc.text(student.lastName || "", x + 2.5, y);
            x += CW[2];
            doc.text(student.firstName || "", x + 2.5, y);
            x += CW[3];
            doc.text(student.className || "", x + 2.5, y);
            doc.setDrawColor(216, 230, 218);
            doc.line(ML, y + 2.2, ML + TW, y + 2.2);
            y += ROW_H;
          };

          const projektTitel = record.number !== 9999
            ? `${record.number}. ${record.name}`
            : record.name;

          const umbruch = () => {
            doc.addPage();
            y = 26;
            titelBand(`${projektTitel} (Forts.)`);
            kopfzeile();
          };

          titelBand(projektTitel);
          kopfzeile();

          const hatSchueler = SLOTS.some((s) => record.slots[s].length > 0);
          if (!hatSchueler) {
            doc.setFontSize(9.5);
            doc.setTextColor(100, 118, 105);
            doc.text("Keine Schüler zugeteilt.", ML + 4, y + 4);
            return;
          }

          for (const slot of SLOTS) {
            const students = record.slots[slot];
            if (!students.length) { continue; }
            if (y + 8 + ROW_H * 2 > PH - 14) { umbruch(); }
            slotBand(slot, students.length);
            for (const student of students) {
              if (y + ROW_H > PH - 14) { umbruch(); }
              schuelerZeile(student);
            }
          }
        });

        applyPdfDecorations(doc, "GAT - Projektliste");
        doc.save("GAT_Zuteilung_nach_Projekt.pdf");
        showMessage("PDF nach Projekt wurde erstellt.", "ok");
      }

      function exportClassPdf() {
        const JsPdf = ensureJsPdf();
        const classes = [...new Set(state.students.map((student) => student.className).filter(Boolean))]
          .sort((a, b) => classSortValue(a) - classSortValue(b));

        if (!classes.length) {
          showMessage("Keine Klassen für PDF vorhanden.", "warn");
          return;
        }

        const doc = new JsPdf({ unit: "mm", format: "a4" });
        const PH = doc.internal.pageSize.getHeight();
        const ML = 14;

        // Spaltenbreiten: Nr. | Name | Projekt / Bemerkung | Slot
        const CW = [11, 62, 88, 17];
        const TW = CW.reduce((s, w) => s + w, 0);
        const ROW_H = 7;

        const clip = (text, maxW) => {
          if (doc.getTextWidth(text) <= maxW) { return text; }
          while (text.length > 4 && doc.getTextWidth(`${text}\u2026`) > maxW) {
            text = text.slice(0, -1);
          }
          return `${text}\u2026`;
        };

        classes.forEach((className, clsIdx) => {
          if (clsIdx > 0) { doc.addPage(); }
          let y = 26;
          let rowNr = 0;

          const titelBand = (label) => {
            doc.setFillColor(219, 240, 224);
            doc.rect(ML, y - 6.5, TW, 10, "F");
            doc.setFontSize(13);
            doc.setFont("helvetica", "bold");
            doc.setTextColor(22, 52, 35);
            doc.text(label, ML + 3, y);
            doc.setFont("helvetica", "normal");
            y += 8;
          };

          const kopfzeile = () => {
            doc.setFillColor(195, 222, 200);
            doc.rect(ML, y - 5.5, TW, 8, "F");
            doc.setFontSize(8.5);
            doc.setFont("helvetica", "bold");
            doc.setTextColor(28, 56, 36);
            let x = ML;
            ["Nr.", "Name", "Projekt / Bemerkung", "Slot"].forEach((l, i) => {
              doc.text(l, x + 2.5, y);
              x += CW[i];
            });
            doc.setFont("helvetica", "normal");
            doc.setDrawColor(148, 188, 156);
            doc.line(ML, y + 2, ML + TW, y + 2);
            y += 8;
          };

          const trennBand = (label) => {
            doc.setFillColor(245, 245, 240);
            doc.rect(ML, y - 4.5, TW, 7, "F");
            doc.setFontSize(8.5);
            doc.setFont("helvetica", "bold");
            doc.setTextColor(110, 110, 100);
            doc.text(label, ML + 4, y);
            doc.setFont("helvetica", "normal");
            y += 7;
          };

          const zeile = (name, projekt, slot, isSpecial) => {
            rowNr++;
            if (rowNr % 2 === 0) {
              doc.setFillColor(250, 254, 251);
              doc.rect(ML, y - 4.5, TW, ROW_H, "F");
            }
            doc.setFontSize(9);
            doc.setTextColor(...(isSpecial ? [112, 108, 88] : [36, 48, 40]));
            let x = ML;
            doc.text(String(rowNr), x + 2.5, y);
            x += CW[0];
            doc.text(name, x + 2.5, y);
            x += CW[1];
            doc.text(clip(projekt, CW[2] - 5), x + 2.5, y);
            x += CW[2];
            doc.text(slot, x + 2.5, y);
            doc.setDrawColor(216, 230, 218);
            doc.line(ML, y + 2.2, ML + TW, y + 2.2);
            y += ROW_H;
          };

          const umbruch = () => {
            doc.addPage();
            y = 26;
            titelBand(`Klasse ${className} (Forts.)`);
            kopfzeile();
          };

          titelBand(`Klasse ${className}`);
          kopfzeile();

          const classStudents = state.students
            .filter((student) => student.className === className)
            .sort((a, b) => a.lastName.localeCompare(b.lastName, "de", { sensitivity: "base" }) ||
                            a.firstName.localeCompare(b.firstName, "de", { sensitivity: "base" }));

          const activeRows = [];
          const absentRows = [];

          for (const student of classStudents) {
            const name = `${student.lastName}, ${student.firstName}`;
            if (student.absent) {
              absentRows.push({ name, projekt: "Abwesend", slot: "—", special: true });
              continue;
            }
            if (student.agMember) {
              const pName = student.agProjectName.trim() || "AG ohne Projekt";
              activeRows.push({ name, projekt: `AG: ${pName}`, slot: "individuell", special: false });
              continue;
            }
            const asgn = state.assignments[student.id];
            if (!asgn) {
              activeRows.push({ name, projekt: "Nicht zugeteilt", slot: "—", special: true });
              continue;
            }
            const project = getProjectById(asgn.projectId);
            activeRows.push({
              name,
              projekt: project ? `${project.number}. ${project.name}` : "Unbekannt",
              slot: asgn.slot,
              special: false
            });
          }

          if (!activeRows.length && !absentRows.length) {
            doc.setFontSize(9.5);
            doc.setTextColor(100, 118, 105);
            doc.text("Keine Schüler vorhanden.", ML + 4, y + 4);
            return;
          }

          for (const row of activeRows) {
            if (y + ROW_H > PH - 14) { umbruch(); }
            zeile(row.name, row.projekt, row.slot, row.special);
          }

          if (absentRows.length) {
            if (y + 7 + ROW_H > PH - 14) { umbruch(); }
            trennBand("Abwesend");
            for (const row of absentRows) {
              if (y + ROW_H > PH - 14) { umbruch(); }
              zeile(row.name, row.projekt, row.slot, row.special);
            }
          }
        });

        applyPdfDecorations(doc, "GAT - Klassenliste");
        doc.save("GAT_Zuteilung_nach_Klasse.pdf");
        showMessage("PDF nach Klasse wurde erstellt.", "ok");
      }

      function handleCsvFile(file) {
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const bytes = new Uint8Array(e.target.result);
            let text = new TextDecoder("utf-8").decode(bytes);
            if (text.includes("\uFFFD")) text = new TextDecoder("windows-1252").decode(bytes);
            importStudentsFromCsv(text);
            renderAll();
          } catch (error) {
            showMessage(error.message || "CSV konnte nicht gelesen werden.", "err");
          }
        };
        reader.onerror = () => showMessage("CSV-Datei konnte nicht geöffnet werden.", "err");
        reader.readAsArrayBuffer(file);
      }

      function importStudentsFromXlsx(file) {
        if (!window.XLSX) {
          showMessage("SheetJS-Bibliothek nicht geladen. Bitte Internetverbindung prüfen.", "err");
          return;
        }
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const wb = window.XLSX.read(new Uint8Array(e.target.result), { type: "array" });
            const ws = wb.Sheets[wb.SheetNames[0]];
            const csvText = window.XLSX.utils.sheet_to_csv(ws, { FS: ";" });
            importStudentsFromCsv(csvText);
            renderAll();
          } catch (error) {
            showMessage(error.message || "XLSX konnte nicht gelesen werden.", "err");
          }
        };
        reader.onerror = () => showMessage("XLSX-Datei konnte nicht geöffnet werden.", "err");
        reader.readAsArrayBuffer(file);
      }

      function handleStudentFile(file) {
        const name = (file.name || "").toLowerCase();
        if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
          importStudentsFromXlsx(file);
        } else {
          handleCsvFile(file);
        }
      }

      function exportXlsx() {
        if (!window.XLSX) {
          showMessage("SheetJS-Bibliothek nicht geladen.", "err");
          return;
        }
        const wb = window.XLSX.utils.book_new();
        const slotMap = buildProjectSlotMap();

        // Sheet 1: Nach Projekt
        const projRows = [["Projekt-Nr.", "Projektname", "Slot", "Nachname", "Vorname", "Klasse"]];
        const sortedProjects = [...state.projects].sort((a, b) => a.number - b.number);
        for (const project of sortedProjects) {
          for (const slot of SLOTS) {
            const students = (slotMap[project.id]?.[slot] || []).sort(compareStudentsByClassLast);
            for (const student of students) {
              projRows.push([project.number, project.name, slot, student.lastName, student.firstName, student.className]);
            }
          }
        }
        window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.aoa_to_sheet(projRows), "Nach Projekt");

        // Sheet 2: Nach Klasse
        const classes = [...new Set(state.students.map((s) => s.className).filter(Boolean))].sort((a, b) => classSortValue(a) - classSortValue(b));
        const classRows = [["Klasse", "Lfd. Nr.", "Nachname", "Vorname", "Projekt", "Slot", "Bemerkung"]];
        let nr = 0;
        for (const className of classes) {
          const students = state.students.filter((s) => s.className === className && !s.absent && !s.agMember).sort(compareStudentsByClassLast);
          for (const student of students) {
            nr += 1;
            const asgn = state.assignments[student.id];
            const project = asgn ? getProjectById(asgn.projectId) : null;
            classRows.push([className, nr, student.lastName, student.firstName, project?.name || "–", asgn?.slot || "–", ""]);
          }
        }
        window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.aoa_to_sheet(classRows), "Nach Klasse");

        // Sheet 3: AG-Listen
        const agRows = [["AG / Liste", "Typ", "Nachname", "Vorname", "Klasse", "Slots"]];
        for (const list of state.agLists) {
          const members = state.students.filter((s) => s.agMember && s.agProjectName === list.name).sort(compareStudentsByClassLast);
          for (const student of members) {
            agRows.push([list.name, list.isSpecial ? "Sonder" : "AG", student.lastName, student.firstName, student.className, getSelectedAgSlots(student).join(", ")]);
          }
        }
        window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.aoa_to_sheet(agRows), "AG-Listen");

        window.XLSX.writeFile(wb, "GAT_Zuteilung.xlsx");
        showMessage("XLSX-Datei wurde exportiert.", "ok");
      }

      // ── localStorage Auto-Save ─────────────────────────────────────
      function saveState() {
        const payload = {
          version: 2,
          savedAt: Date.now(),
          students: state.students,
          projects: state.projects,
          assignments: state.assignments,
          agProjects: state.agProjects,
          agLists: state.agLists,
          agExtra: state.agExtra,
          slotPriorities: state.slotPriorities,
          pdf: state.pdf
        };
        try {
          localStorage.setItem(LS_KEY, JSON.stringify(payload));
        } catch (_) {
          // Quota-Fehler ignorieren
        }
      }

      function loadSavedState() {
        const raw = localStorage.getItem(LS_KEY);
        if (!raw) {
          return false;
        }
        try {
          loadJsonSnapshot(raw);
          return true;
        } catch (_) {
          return false;
        }
      }

      function handleJsonFile(file) {
        const reader = new FileReader();
        reader.onload = () => {
          try {
            loadJsonSnapshot(String(reader.result || ""));
            renderAll();
            showMessage("JSON-Datei wurde geladen.", "ok");
          } catch (error) {
            showMessage(error.message || "JSON konnte nicht geladen werden.", "err");
          }
        };
        reader.onerror = () => showMessage("JSON-Datei konnte nicht geöffnet werden.", "err");
        reader.readAsText(file, "utf-8");
      }

      function handleDrop(event) {
        const zone = event.target.closest("[data-dropzone='true']");
        if (!zone) {
          return;
        }
        event.preventDefault();
        zone.classList.remove("over");
        const studentId = event.dataTransfer.getData("text/plain");
        if (!studentId) {
          return;
        }

        // AG-Listen-Pool (zurück in den Pool ziehen)
        if (zone.hasAttribute("data-aglist-pool")) {
          removeStudentFromAgList(studentId);
          return;
        }

        // AG-Listen-Karte (aus Pool oder anderer Karte in eine AG/Sonderliste ziehen)
        if (zone.hasAttribute("data-aglist-id")) {
          addStudentToAgList(studentId, zone.dataset.aglistId);
          return;
        }

        // Sonderliste Slot (Sonstiges-Tab: Schüler zwischen Slot-Spalten verschieben)
        if (zone.hasAttribute("data-sonder-list-id")) {
          moveSonderStudentToSlot(studentId, zone.dataset.sonderListId, zone.dataset.targetSlot);
          return;
        }

        // AG-Board Dropzone (Step-4 AG-Verwaltung)
        if (zone.hasAttribute("data-ag-project")) {
          const student = getStudentById(studentId);
          if (!student?.agMember) {
            showMessage("Nur AG-Mitglieder können in die AG-Verwaltung verschoben werden.", "warn");
            return;
          }
          moveStudentToAg(studentId, zone.dataset.agProject || "");
          return;
        }

        // Regulärer Board-Drop
        const student = getStudentById(studentId);
        if (student?.agMember && !isSonderStudent(student)) {
          showMessage("AG-Mitglieder können nicht der regulären Zuteilung hinzugefuegt werden.", "warn");
          return;
        }

        const targetProjectId = zone.dataset.targetProjectId || "";
        const targetSlot = zone.dataset.targetSlot || "";

        // Sonder-Validierung: Sonder-agMember dürfen nur in Sonderprojekte
        if (targetProjectId && isSonderStudent(student)) {
          const targetProject = getProjectById(targetProjectId);
          if (!targetProject?.isSpecial) {
            showMessage("Sonder-Schüler können nur in Sonderprojekte eingetragen werden.", "warn");
            return;
          }
        }

        moveStudent(studentId, targetProjectId || null, targetSlot || null);
      }

      function wireEvents() {
        document.addEventListener("click", (event) => {
          const target = event.target.closest("[data-action]");
          if (!target) {
            return;
          }

          const action = target.dataset.action;

          if (action === "goto-step") {
            state.activeStep = safeInt(target.dataset.step) || 1;
            hideMessage();
            renderStepNavigation();
            return;
          }

          if (action === "add-project") {
            state.projects.push(createProject());
            renderStep2();
            renderStep3();
            renderStep4();
            return;
          }

          if (action === "load-samples") {
            if (state.projects.length && !window.confirm("Aktuelle Projekte werden ersetzt. Fortfahren?")) {
              return;
            }
            loadSampleProjects();
            renderAll();
            return;
          }

          if (action === "reset-projects") {
            if (!window.confirm("Alle Projekte wirklich löschen?")) {
              return;
            }
            state.projects = [];
            state.assignments = {};
            state.conflicts = [];
            showMessage("Alle Projekte wurden geloescht.", "warn");
            renderAll();
            return;
          }

          if (action === "delete-project") {
            const projectId = target.dataset.projectId;
            state.projects = state.projects.filter((project) => project.id !== projectId);
            sanitizeAssignments();
            showMessage("Projekt wurde geloescht.", "warn");
            renderAll();
            return;
          }

          if (action === "generate") {
            executeGeneration();
            return;
          }

          if (action === "close-conflict-modal") {
            closeConflictModal();
            return;
          }

          if (action === "confirm-conflict-modal") {
            if (state.modal.mode === "pre") {
              closeConflictModal();
              const result = runAllocation();
              if (result.conflicts.length) {
                openConflictModal(
                  "post",
                  "Konflikte nach Algorithmuslauf",
                  "Nicht alle Bedarfe konnten vollständig gedeckt werden.",
                  result.conflicts,
                  result
                );
                return;
              }
              state.assignments = result.assignments;
              state.conflicts = result.conflicts;
              showMessage("Zuteilung wurde trotz Warnung erzeugt.", "warn");
              renderAll();
              return;
            }

            if (state.modal.mode === "post" && state.modal.pendingResult) {
              state.assignments = deepClone(state.modal.pendingResult.assignments);
              state.conflicts = [...state.modal.pendingResult.conflicts];
              closeConflictModal();
              showMessage("Zuteilung mit Konflikten wurde gespeichert.", "warn");
              renderAll();
            }
            return;
          }

          if (action === "regenerate") {
            const preConflicts = computePreConflicts();
            if (preConflicts.length) {
              openConflictModal(
                "pre",
                "Unterbesetzungs-Warnung",
                "Mindestens eine Slot+Stufe-Kombination ist voraussichtlich unterbesetzt.",
                preConflicts
              );
              return;
            }
            finalizeGeneration();
            return;
          }

          if (action === "save-json") {
            downloadJsonSnapshot();
            return;
          }

          if (action === "open-move") {
            const studentId = target.dataset.studentId;
            if (studentId) {
              openMoveDialog(studentId);
            }
            return;
          }

          if (action === "close-move-modal") {
            closeMoveDialog();
            return;
          }

          if (action === "confirm-move") {
            const studentId = state.moveDialog.studentId;
            if (!studentId) {
              closeMoveDialog();
              return;
            }
            const targetProjectId = byId("move-project-select").value || null;
            const targetSlot = byId("move-slot-select").value || "S1";
            moveStudent(studentId, targetProjectId, targetProjectId ? targetSlot : null);
            closeMoveDialog();
            return;
          }

          if (action === "export-project-pdf") {
            try {
              exportProjectPdf();
            } catch (error) {
              showMessage(error.message || "Projekt-PDF konnte nicht erstellt werden.", "err");
            }
            return;
          }

          if (action === "export-class-pdf") {
            try {
              exportClassPdf();
            } catch (error) {
              showMessage(error.message || "Klassen-PDF konnte nicht erstellt werden.", "err");
            }
          }

          // ── Theme Toggle ─────────────────────────────────────────────────

          if (action === "toggle-theme") {
            const html = document.documentElement;
            const next = (html.dataset.theme || "light") === "light" ? "dark" : "light";
            html.classList.add("theme-transitioning");
            html.dataset.theme = next;
            try { localStorage.setItem("gat_theme", next); } catch (_) {}
            setTimeout(() => html.classList.remove("theme-transitioning"), 350);
            return;
          }

          // ── AG-Board-Aktionen ────────────────────────────────────────────

          if (action === "switch-board") {
            switchBoard(target.dataset.board || "regular");
            return;
          }

          if (action === "reset-step4-filter") {
            const s = byId("step4-search");
            const g = byId("step4-grade-filter");
            const c = byId("step4-class-filter");
            if (s) s.value = "";
            if (g) g.value = "all";
            if (c) c.value = "all";
            applyStep4Filter();
            return;
          }

          if (action === "add-ag-project") {
            const name = window.prompt("Name des neuen AG-Projekts:");
            if (!name?.trim()) {
              return;
            }
            const trimmed = name.trim();
            if (!state.agProjects.includes(trimmed)) {
              state.agProjects.push(trimmed);
            }
            renderAgBoard();
            return;
          }

          if (action === "delete-ag-project") {
            const agName = target.dataset.agName || "";
            const affected = state.students.filter(
              (student) => student.agMember && student.agProjectName.trim() === agName
            );
            const msg = affected.length
              ? `AG "${agName}" entfernen? ${affected.length} Schüler bleiben AG-Mitglieder, verlieren aber die AG-Zuteilung.`
              : `Leere AG "${agName}" entfernen?`;
            if (!window.confirm(msg)) {
              return;
            }
            // Schüler-Zuordnung löschen, AG aus Liste entfernen
            for (const student of affected) {
              student.agProjectName = "";
            }
            state.agProjects = state.agProjects.filter((name) => name !== agName);
            renderAgBoard();
            renderStep1();
            renderStep5();
            return;
          }

          if (action === "open-ag-edit") {
            openAgEditDialog(target.dataset.studentId || "");
            return;
          }

          if (action === "close-ag-edit-modal") {
            closeAgEditDialog();
            return;
          }

          if (action === "confirm-ag-edit") {
            confirmAgEdit();
            return;
          }

          if (action === "remove-from-ag") {
            const student = getStudentById(state.agEditDialog.studentId);
            if (student) {
              student.agMember = false;
              student.agProjectName = "";
              student.agSlots = emptyAgSlots();
            }
            closeAgEditDialog();
            renderAll();
            return;
          }

          // ── Step-1-Sub-Tabs ──────────────────────────────────────────────

          if (action === "step1-tab") {
            state.activeStep1Tab = target.dataset.tab || "students";
            document.querySelectorAll(".step1-tab").forEach((btn) => {
              btn.classList.toggle("active", btn.dataset.tab === state.activeStep1Tab);
            });
            document.querySelectorAll(".step1-section").forEach((sec) => {
              sec.classList.toggle("active", sec.id === `step1-${state.activeStep1Tab}`);
            });
            return;
          }

          // ── XLSX Export ──────────────────────────────────────────────────

          if (action === "export-xlsx") {
            try {
              exportXlsx();
            } catch (error) {
              showMessage(error.message || "XLSX-Export fehlgeschlagen.", "err");
            }
            return;
          }

          // ── localStorage löschen ────────────────────────────────────────

          if (action === "clear-storage") {
            if (!window.confirm("Gespeicherten Browser-Zustand wirklich löschen?")) {
              return;
            }
            try { localStorage.removeItem(LS_KEY); } catch (_) {}
            showMessage("Browser-Speicher wurde geloescht.", "warn");
            return;
          }

          // ── AG-Listen (Step-1) ───────────────────────────────────────────

          if (action === "add-ag-list") {
            const name = window.prompt("Name der neuen AG-Gruppe:");
            if (!name?.trim()) return;
            state.agLists.push(createAgList(name.trim(), false));
            renderStep1AgLists();
            saveState();
            return;
          }

          if (action === "add-sonder-list") {
            const name = window.prompt("Name der neuen Sonderliste:");
            if (!name?.trim()) return;
            state.agLists.push(createAgList(name.trim(), true));
            renderStep1AgLists();
            saveState();
            return;
          }

          if (action === "delete-aglist") {
            const listId = target.dataset.aglistId || "";
            const list = state.agLists.find((l) => l.id === listId);
            if (!list) return;
            const affected = state.students.filter(
              (s) => s.agMember && s.agProjectName === list.name
            );
            const msg = affected.length
              ? `"${list.name}" löschen? ${affected.length} Schüler werden aus der Liste entfernt.`
              : `"${list.name}" löschen?`;
            if (!window.confirm(msg)) return;
            for (const s of affected) {
              s.agMember = false;
              s.agProjectName = "";
              s.agSlots = emptyAgSlots();
            }
            state.agLists = state.agLists.filter((l) => l.id !== listId);
            delete state.agExtra[listId];
            renderAll();
            return;
          }

          if (action === "remove-from-aglist") {
            const studentId = target.dataset.studentId || "";
            removeStudentFromAgList(studentId);
            renderAll();
            return;
          }

          if (action === "remove-from-ag-extra") {
            removeStudentFromAgExtra(target.dataset.studentId || "", target.dataset.aglistId || "");
            return;
          }

          if (action === "ag-ctx-add") {
            addStudentToAgExtra(target.dataset.studentId || "", target.dataset.aglistId || "");
            byId("ag-ctx-menu").style.display = "none";
            return;
          }

          // ── Sonderprojekt-Toggle ─────────────────────────────────────────

          if (action === "toggle-special") {
            const project = getProjectById(target.dataset.projectId);
            if (!project) return;
            project.isSpecial = target.checked;
            renderStep3();
            renderStep4();
            saveState();
            return;
          }
        });

        document.addEventListener("input", (event) => {
          const target = event.target;
          const action = target.dataset.action;

          if (target.id === "aglist-pool-search") {
            const q = target.value.toLowerCase();
            byId("step1-aglist-board")?.querySelectorAll(".aglist-pool-zone .pool-chip").forEach((chip) => {
              chip.style.display = chip.textContent.toLowerCase().includes(q) ? "" : "none";
            });
            return;
          }

          if (target.id === "student-search") {
            state.filters.search = target.value;
            renderStep1();
            return;
          }

          if (target.id === "step4-search") {
            applyStep4Filter();
            return;
          }

          if (action === "project-name") {
            const project = getProjectById(target.dataset.projectId);
            if (!project) {
              return;
            }
            // Kein .trim() hier — entfernt führende Leerzeichen beim Tippen.
            // Kein renderStep2() — würde das fokussierte Feld zerstören.
            project.name = target.value || "Projekt";
            renderStep3();
            renderStep4();
            return;
          }

          if (action === "project-number") {
            const project = getProjectById(target.dataset.projectId);
            if (!project) {
              return;
            }
            project.number = Math.max(1, safeInt(target.value) || 1);
            // Kein renderStep2() — würde das fokussierte Feld zerstören.
            renderDemandSummary();
            renderStep3();
            renderStep4();
            return;
          }

          if (action === "project-demand") {
            const project = getProjectById(target.dataset.projectId);
            if (!project) {
              return;
            }
            const grade = safeInt(target.dataset.grade);
            const slot = target.dataset.slot;
            if (!GRADES.includes(grade) || !SLOTS.includes(slot)) {
              return;
            }
            project.demands[grade][slot] = safeInt(target.value);
            // Nur Zusammenfassung aktualisieren, nicht den gesamten Projektbereich.
            renderDemandSummary();
            renderStep3();
            return;
          }

          if (action === "pref-class-weight") {
            const project = getProjectById(target.dataset.projectId);
            if (!project) {
              return;
            }
            const className = target.dataset.className;
            if (!className || !ALL_CLASSES.includes(className)) {
              return;
            }
            if (!Object.prototype.hasOwnProperty.call(project.preference.classes, className)) {
              return;
            }
            project.preference.classes[className] = Math.max(1, safeInt(target.value) || 1);
            renderDemandSummary();
            return;
          }

          if (action === "ag-project") {
            const student = getStudentById(target.dataset.studentId);
            if (!student) {
              return;
            }
            student.agProjectName = target.value;
            // Kein renderStep1() — würde das fokussierte Feld zerstören.
            updateTopStats();
            renderStep5();
            return;
          }

          if (target.id === "school-header") {
            state.pdf.schoolHeader = target.value;
            return;
          }
        });

        document.addEventListener("change", (event) => {
          const target = event.target;
          const action = target.dataset.action;

          if (target.id === "class-filter") {
            state.filters.className = target.value;
            renderStep1();
            return;
          }

          if (target.id === "grade-filter") {
            state.filters.grade = target.value;
            renderStep1();
            return;
          }

          if (target.id === "step4-grade-filter" || target.id === "step4-class-filter") {
            applyStep4Filter();
            return;
          }

          if (target.id === "csv-input" && target.files?.[0]) {
            handleStudentFile(target.files[0]);
            target.value = "";
            return;
          }

          if (target.id === "json-input" && target.files?.[0]) {
            handleJsonFile(target.files[0]);
            target.value = "";
            return;
          }

          if (action === "toggle-absent") {
            const student = getStudentById(target.dataset.studentId);
            if (!student) {
              return;
            }
            student.absent = target.checked;
            if (student.absent) {
              delete state.assignments[student.id];
            }
            renderAll();
            return;
          }

          if (action === "toggle-ag") {
            const student = getStudentById(target.dataset.studentId);
            if (!student) {
              return;
            }
            student.agMember = target.checked;
            if (student.agMember) {
              if (!student.agProjectName) {
                student.agProjectName = "AG ohne Projekt";
              }
              if (!getSelectedAgSlots(student).length) {
                student.agSlots.S1 = true;
              }
              delete state.assignments[student.id];
            } else {
              student.agProjectName = "";
              student.agSlots = emptyAgSlots();
            }
            renderAll();
            return;
          }

          if (action === "ag-slot") {
            const student = getStudentById(target.dataset.studentId);
            if (!student || !student.agMember) {
              return;
            }
            const slot = target.dataset.slot;
            if (!SLOTS.includes(slot)) {
              return;
            }
            student.agSlots[slot] = target.checked;
            if (!getSelectedAgSlots(student).length) {
              student.agSlots.S1 = true;
              showMessage("Mindestens ein AG-Slot muss aktiv bleiben — S1 wurde beibehalten.", "warn");
            }
            renderStep1();
            renderStep5();
            return;
          }

          // Beim Verlassen (blur) des Projektnamen-Felds: Trim anwenden und Step2 aktualisieren.
          if (action === "project-name") {
            const project = getProjectById(target.dataset.projectId);
            if (project) {
              project.name = target.value.trim() || "Projekt";
              renderStep2();
              renderStep3();
              renderStep4();
            }
            return;
          }

          // Beim Verlassen des AG-Projektfelds: Step1 und Step5 aktualisieren.
          if (action === "ag-project") {
            const student = getStudentById(target.dataset.studentId);
            if (student) {
              student.agProjectName = target.value.trim();
              renderStep1();
              renderStep5();
            }
            return;
          }

          if (action === "pref-mode") {
            const project = getProjectById(target.dataset.projectId);
            if (!project) {
              return;
            }
            project.preference.mode = target.value === "specific" ? "specific" : "all";
            renderStep2();
            return;
          }

          if (action === "pref-class-toggle") {
            const project = getProjectById(target.dataset.projectId);
            const className = target.dataset.className;
            if (!project || !ALL_CLASSES.includes(className)) {
              return;
            }
            if (target.checked) {
              project.preference.classes[className] = project.preference.classes[className] || 25;
            } else {
              delete project.preference.classes[className];
            }
            renderStep2();
            return;
          }

          if (target.id === "opt-header") {
            state.pdf.showHeader = target.checked;
            return;
          }

          if (target.id === "opt-date") {
            state.pdf.showDate = target.checked;
            return;
          }

          if (target.id === "opt-pages") {
            state.pdf.showPageNumbers = target.checked;
            return;
          }

          if (target.id === "project-order") {
            state.pdf.projectOrder = target.value === "name" ? "name" : "id";
          }

          if (target.dataset.priorityGrade) {
            const grade = Number(target.dataset.priorityGrade);
            const slot = target.dataset.prioritySlot;
            if (!state.slotPriorities[grade]) {
              state.slotPriorities[grade] = {};
            }
            if (target.value) {
              state.slotPriorities[grade][slot] = target.value;
            } else {
              delete state.slotPriorities[grade][slot];
            }
            saveState();
            return;
          }
        });

        // dragstart auf document registrieren, damit auch Chips aus der Übrig-Zone gezogen werden können.
        document.addEventListener("dragstart", (event) => {
          const chip = event.target.closest(".student-chip, .pool-chip, .sonder-chip");
          if (!chip) {
            return;
          }
          event.dataTransfer.effectAllowed = "move";
          event.dataTransfer.setData("text/plain", chip.dataset.studentId || "");
        });

        document.addEventListener("dragover", (event) => {
          const zone = event.target.closest("[data-dropzone='true'], .ag-dropzone, .aglist-dropzone");
          if (!zone) {
            return;
          }
          event.preventDefault();
          zone.classList.add("over");
        });

        document.addEventListener("dragleave", (event) => {
          const zone = event.target.closest("[data-dropzone='true'], .ag-dropzone, .aglist-dropzone");
          if (zone) {
            zone.classList.remove("over");
          }
        });

        document.addEventListener("drop", handleDrop);

        byId("move-project-select").addEventListener("change", () => {
          const hasProject = Boolean(byId("move-project-select").value);
          byId("move-slot-select").disabled = !hasProject;
        });

        byId("btn-csv-trigger").addEventListener("click", () => byId("csv-input").click());
        byId("btn-json-trigger").addEventListener("click", () => byId("json-input").click());

        // ── AG-Duplikat-Kontextmenü ─────────────────────────────────────────
        const ctxMenu = document.createElement("div");
        ctxMenu.id = "ag-ctx-menu";
        ctxMenu.className = "ag-ctx-menu";
        ctxMenu.style.display = "none";
        document.body.appendChild(ctxMenu);

        document.addEventListener("contextmenu", (e) => {
          const chip = e.target.closest(".aglist-dropzone [data-student-id]");
          if (!chip) { ctxMenu.style.display = "none"; return; }
          e.preventDefault();
          const studentId = chip.dataset.studentId;
          const student = getStudentById(studentId);
          if (!student) return;
          const alreadyIn = new Set([
            state.agLists.find((l) => l.name === student.agProjectName)?.id,
            ...Object.entries(state.agExtra).filter(([, ids]) => ids.includes(studentId)).map(([lid]) => lid)
          ].filter(Boolean));
          const targets = state.agLists.filter((l) => !alreadyIn.has(l.id));
          if (!targets.length) {
            showMessage("Schüler ist bereits in allen Listen.", "info");
            return;
          }
          ctxMenu.innerHTML = `
            <div class="ag-ctx-title">Auch hinzufügen zu:</div>
            ${targets.map((l) => `
              <button type="button" data-action="ag-ctx-add"
                data-student-id="${escapeHtml(studentId)}"
                data-aglist-id="${escapeHtml(l.id)}">
                ${escapeHtml(l.name)}
              </button>
            `).join("")}
          `;
          ctxMenu.style.left = `${e.clientX + 2}px`;
          ctxMenu.style.top = `${e.clientY + 2}px`;
          ctxMenu.style.display = "block";
          const rect = ctxMenu.getBoundingClientRect();
          if (rect.right > window.innerWidth) ctxMenu.style.left = `${e.clientX - rect.width - 2}px`;
          if (rect.bottom > window.innerHeight) ctxMenu.style.top = `${e.clientY - rect.height - 2}px`;
        });

        document.addEventListener("pointerdown", (e) => {
          if (!e.target.closest("#ag-ctx-menu")) ctxMenu.style.display = "none";
        });
      }

      function bootstrap() {
        wireEvents();
        const restored = loadSavedState();
        if (restored) {
          // agListIdCounter hinter dem höchsten geladenen ID setzen
          state.agLists.forEach((l) => {
            const n = parseInt(l.id.replace(/\D/g, ""), 10);
            if (!isNaN(n) && n >= agListIdCounter) agListIdCounter = n + 1;
          });
          showMessage("Gespeicherter Zustand wurde wiederhergestellt.", "info");
        } else {
          state.projects.push(createProject(1, "Beispielprojekt"));
        }
        renderAll();
        hideMessage();
      }

      bootstrap();
    })();
