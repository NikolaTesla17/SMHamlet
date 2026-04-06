const MAX_SCHEDULE_SOLUTIONS = 50;
const STORAGE_KEY_REHEARSALS = 'rehearsal_scheduler_rehearsals_v1';

let responsesFileName = '';
let responsesHeaders = [];
let responsesRows = [];
let people = [];
let rehearsals = [];
let generatedSchedules = [];
let filteredScheduleIndexes = [];

const elements = {};

document.addEventListener('DOMContentLoaded', () => {
  cacheElements();
  bindEvents();
  loadRehearsalsFromStorage();
  renderAll();
});

function cacheElements() {
  [
    'dropZone', 'responsesFile', 'loadedFileName', 'peopleCount', 'availabilityCount',
    'title', 'hours', 'peopleList', 'mustFollow', 'editingId', 'saveButton', 'clearButton',
    'status', 'rehearsalList', 'rehearsalCount', 'generateSchedulesBtn', 'schedulePicker',
    'scheduleCount', 'scheduleWindow', 'filterDay', 'filterStart', 'filterEnd', 'filterSummary',
    'clearSearchBtn', 'exportRehearsalsBtn', 'importRehearsalsFile'
  ].forEach(id => {
    elements[id] = document.getElementById(id);
  });
}

function bindEvents() {
  elements.responsesFile.addEventListener('change', event => {
    const file = event.target.files?.[0];
    if (file) {
      handleResponsesFile(file);
    }
  });

  ['dragenter', 'dragover'].forEach(eventName => {
    elements.dropZone.addEventListener(eventName, event => {
      event.preventDefault();
      elements.dropZone.classList.add('drag-over');
    });
  });

  ['dragleave', 'drop'].forEach(eventName => {
    elements.dropZone.addEventListener(eventName, event => {
      event.preventDefault();
      elements.dropZone.classList.remove('drag-over');
    });
  });

  elements.dropZone.addEventListener('drop', event => {
    const file = event.dataTransfer?.files?.[0];
    if (file) {
      handleResponsesFile(file);
    }
  });

  elements.dropZone.addEventListener('keydown', event => {
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      elements.responsesFile.click();
    }
  });

  elements.saveButton.addEventListener('click', submitRehearsalForm);
  elements.clearButton.addEventListener('click', resetForm);
  elements.generateSchedulesBtn.addEventListener('click', generateSchedules);
  elements.schedulePicker.addEventListener('change', renderSelectedSchedule);
  elements.filterDay.addEventListener('change', applyScheduleFilters);
  elements.filterStart.addEventListener('change', applyScheduleFilters);
  elements.filterEnd.addEventListener('change', applyScheduleFilters);
  elements.clearSearchBtn.addEventListener('click', clearScheduleFilters);
  elements.exportRehearsalsBtn.addEventListener('click', exportRehearsalsJson);
  elements.importRehearsalsFile.addEventListener('change', event => {
    const file = event.target.files?.[0];
    if (file) {
      importRehearsalsJson(file);
    }
  });
}

async function handleResponsesFile(file) {
  try {
    setStatus('Loading CSV...', '');
    const parsed = await parseCsvFile(file);
    const headers = parsed.meta?.fields || [];
    const rows = parsed.data || [];

    if (!headers.length) {
      throw new Error('The CSV appears to have no header row.');
    }

    findNameColumn(headers);

    responsesFileName = file.name;
    responsesHeaders = headers;
    responsesRows = rows;
    people = getPeopleFromResponsesData();

    resetGeneratedSchedules();
    resetForm();
    renderAll();
    setStatus(`Loaded ${file.name}.`, 'success');
  } catch (error) {
    setStatus(error.message || 'Failed to load CSV.', 'error');
  }
  console.log(
    'Availability header names:',
    getAvailabilityDaySpecsFromData().map(spec => spec.header)
  );
}

function parseCsvFile(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      quoteChar: '"',
      escapeChar: '\\',
      complete: results => {
        console.log('Parsed headers:', results.meta?.fields);
        console.log('Parse errors:', results.errors);
        resolve(results);
      },
      error: err => reject(err)
    });
  });
}

function findNameColumn(headers) {
  const candidates = ['Name', 'Full Name', 'Your Name'];
  for (const candidate of candidates) {
    if (headers.includes(candidate)) {
      return candidate;
    }
  }
  throw new Error('Could not find a name column. Expected Name, Full Name, or Your Name.');
}

function getPeopleFromResponsesData() {
  if (!responsesHeaders.length) {
    return [];
  }

  const nameHeader = findNameColumn(responsesHeaders);
  const seen = new Set();
  const list = [];

  responsesRows.forEach(row => {
    const rawName = String(row[nameHeader] || '').trim();
    if (!rawName) {
      return;
    }

    const key = canonicalizeName(rawName);
    if (!seen.has(key)) {
      seen.add(key);
      list.push(rawName);
    }
  });

  return list.sort((a, b) => a.localeCompare(b));
}

function getAvailabilityDaySpecsFromData() {
  const specs = [];

  responsesHeaders.forEach(header => {
    const headerText = String(header || '').trim();
    if (!headerText) {
      return;
    }

    const lower = headerText.toLowerCase();
    if (!lower.includes('when are you available')) {
      return;
    }

    const match = headerText.match(/on\s+([A-Za-z]+),\s+([A-Za-z]+\s+\d+(?:st|nd|rd|th)?)/i);
    specs.push({
      header: headerText,
      dayName: match ? match[1] : `Day ${specs.length + 1}`,
      dateLabel: match ? match[2] : headerText,
      dayIndex: specs.length
    });
  });

  return specs;
}

function buildAvailabilityMapFromData(daySpecs) {
  const nameHeader = findNameColumn(responsesHeaders);
  const map = {};

  responsesRows.forEach(row => {
    const rawName = String(row[nameHeader] || '').trim();
    if (!rawName) {
      return;
    }

    const key = canonicalizeName(rawName);
    if (!map[key]) {
      map[key] = {
        displayName: rawName,
        byDay: {}
      };
    }

    daySpecs.forEach(day => {
      const rawCell = String(row[day.header] || '');
      const labels = extractTimeLabels(rawCell);
      map[key].byDay[day.dayIndex] = new Set(labels);
    });
  });

  return map;
}

function canonicalizeName(name) {
  return String(name || '').trim().toLowerCase();
}

function extractTimeLabels(text) {
  const matches = String(text || '').match(/\b\d{1,2}:\d{2}\s*[ap]m\b/gi);
  if (!matches) {
    return [];
  }

  const unique = [...new Set(matches.map(normalizeTimeLabel))];
  unique.sort((a, b) => timeLabelToMinutes(a) - timeLabelToMinutes(b));
  return unique;
}

function normalizeTimeLabel(label) {
  return String(label || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function timeLabelToMinutes(label) {
  const match = normalizeTimeLabel(label).match(/^(\d{1,2}):(\d{2})\s*([ap]m)$/);
  if (!match) {
    throw new Error(`Could not parse time label "${label}".`);
  }

  let hour = Number(match[1]);
  const minute = Number(match[2]);
  const ampm = match[3];

  if (hour === 12) {
    hour = 0;
  }
  if (ampm === 'pm') {
    hour += 12;
  }

  return hour * 60 + minute;
}

function minutesToTimeLabel(minutes) {
  let total = Number(minutes);
  total = ((total % 1440) + 1440) % 1440;

  let hour24 = Math.floor(total / 60);
  const minute = total % 60;
  const ampm = hour24 >= 12 ? 'pm' : 'am';
  let hour12 = hour24 % 12;
  if (hour12 === 0) {
    hour12 = 12;
  }

  return `${hour12}:${String(minute).padStart(2, '0')} ${ampm}`;
}

function submitRehearsalForm() {
  try {
    const id = elements.editingId.value || crypto.randomUUID();
    const title = elements.title.value.trim();
    const hours = Number(elements.hours.value);
    const mustFollowId = elements.mustFollow.value;
    const neededPeople = getSelectedPeople();

    if (!title) {
      throw new Error('Title is required.');
    }
    if (!Number.isInteger(hours) || hours <= 0) {
      throw new Error('Hours must be a whole number greater than 0.');
    }
    if (!neededPeople.length) {
      throw new Error('Select at least one person.');
    }
    if (mustFollowId && mustFollowId === id) {
      throw new Error('A rehearsal cannot depend on itself.');
    }

    const prior = rehearsals.find(r => r.id === mustFollowId);
    const record = {
      id,
      title,
      hours,
      neededPeople: neededPeople.join(', '),
      mustFollowId,
      mustFollowTitle: prior ? prior.title : ''
    };

    const existingIndex = rehearsals.findIndex(r => r.id === id);
    if (existingIndex >= 0) {
      rehearsals[existingIndex] = record;
      setStatus('Rehearsal updated.', 'success');
    } else {
      rehearsals.push(record);
      setStatus('Rehearsal added.', 'success');
    }

    refreshMustFollowTitles();
    saveRehearsalsToStorage();
    resetGeneratedSchedules();
    resetForm();
    renderAll();
  } catch (error) {
    setStatus(error.message || 'Failed to save rehearsal.', 'error');
  }
}

function getSelectedPeople() {
  return [...document.querySelectorAll('input[name="neededPeople"]:checked')].map(cb => cb.value);
}

function setSelectedPeople(names) {
  const selected = new Set(names);
  document.querySelectorAll('input[name="neededPeople"]').forEach(cb => {
    cb.checked = selected.has(cb.value);
  });
}

function resetForm() {
  elements.editingId.value = '';
  elements.title.value = '';
  elements.hours.value = '';
  elements.mustFollow.value = '';
  elements.saveButton.textContent = 'Add rehearsal';
  setSelectedPeople([]);
  renderMustFollowDropdown();
}

function startEdit(id) {
  const rehearsal = rehearsals.find(r => r.id === id);
  if (!rehearsal) {
    setStatus('Could not find rehearsal to edit.', 'error');
    return;
  }

  elements.editingId.value = rehearsal.id;
  elements.title.value = rehearsal.title;
  elements.hours.value = rehearsal.hours;
  elements.saveButton.textContent = 'Save changes';
  renderMustFollowDropdown(rehearsal.id);
  elements.mustFollow.value = rehearsal.mustFollowId || '';
  setSelectedPeople(splitPeopleList(rehearsal.neededPeople));
  window.scrollTo({ top: 0, behavior: 'smooth' });
  setStatus('Editing rehearsal.', '');
}

function deleteRehearsal(id) {
  const rehearsal = rehearsals.find(r => r.id === id);
  if (!rehearsal) {
    return;
  }

  const ok = window.confirm(`Delete rehearsal "${rehearsal.title}"?`);
  if (!ok) {
    return;
  }

  rehearsals = rehearsals.filter(r => r.id !== id);
  rehearsals.forEach(r => {
    if (r.mustFollowId === id) {
      r.mustFollowId = '';
      r.mustFollowTitle = '';
    }
  });

  saveRehearsalsToStorage();
  resetGeneratedSchedules();
  resetForm();
  renderAll();
  setStatus('Rehearsal deleted.', 'success');
}

function refreshMustFollowTitles() {
  rehearsals.forEach(r => {
    const prior = rehearsals.find(other => other.id === r.mustFollowId);
    r.mustFollowTitle = prior ? prior.title : '';
  });
}

function renderPeopleList() {
  elements.peopleList.innerHTML = '';

  if (!people.length) {
    elements.peopleList.innerHTML = '<div class="muted">Load a responses CSV to populate people.</div>';
    return;
  }

  people.forEach((name, index) => {
    const row = document.createElement('div');
    row.className = 'person-row';

    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.id = `person_${index}`;
    checkbox.name = 'neededPeople';
    checkbox.value = name;

    const label = document.createElement('label');
    label.htmlFor = checkbox.id;
    label.textContent = name;

    row.appendChild(checkbox);
    row.appendChild(label);
    elements.peopleList.appendChild(row);
  });
}

function renderMustFollowDropdown(excludeId = '') {
  elements.mustFollow.innerHTML = '<option value="">None</option>';

  rehearsals
    .filter(r => r.id !== excludeId)
    .forEach(r => {
      const option = document.createElement('option');
      option.value = r.id;
      option.textContent = r.title;
      elements.mustFollow.appendChild(option);
    });
}

function renderRehearsalList() {
  elements.rehearsalList.innerHTML = '';
  elements.rehearsalCount.textContent = `${rehearsals.length} rehearsal${rehearsals.length === 1 ? '' : 's'}`;

  if (!rehearsals.length) {
    elements.rehearsalList.innerHTML = '<div class="muted">No rehearsals yet.</div>';
    return;
  }

  rehearsals.forEach(r => {
    const div = document.createElement('div');
    div.className = 'rehearsal-item';
    div.innerHTML = `
      <div class="rehearsal-title">${escapeHtml(r.title)}</div>
      <div>Hours: ${escapeHtml(r.hours)}</div>
      <div>People needed: ${escapeHtml(r.neededPeople || 'None')}</div>
      <div>Must happen after: ${escapeHtml(r.mustFollowTitle || 'None')}</div>
      <div class="item-actions">
        <button class="button secondary" type="button" data-action="edit">Edit</button>
        <button class="button secondary" type="button" data-action="delete">Delete</button>
      </div>
    `;

    div.querySelector('[data-action="edit"]').addEventListener('click', () => startEdit(r.id));
    div.querySelector('[data-action="delete"]').addEventListener('click', () => deleteRehearsal(r.id));
    elements.rehearsalList.appendChild(div);
  });
}

function renderCsvMeta() {
  const daySpecs = responsesHeaders.length ? getAvailabilityDaySpecsFromData() : [];
  elements.loadedFileName.textContent = responsesFileName || 'None';
  elements.peopleCount.textContent = String(people.length);
  elements.availabilityCount.textContent = String(daySpecs.length);
}

function setStatus(message, type) {
  elements.status.textContent = message || '';
  elements.status.className = 'status';
  if (type) {
    elements.status.classList.add(type);
  }
}

function renderAll() {
  renderCsvMeta();
  renderPeopleList();
  renderMustFollowDropdown(elements.editingId.value);
  renderRehearsalList();
  populateSchedulePicker();
  renderSelectedSchedule();
  updateFilterSummary();
}

function resetGeneratedSchedules() {
  generatedSchedules = [];
  filteredScheduleIndexes = [];
}

function splitPeopleList(value) {
  return String(value || '')
    .split(',')
    .map(x => x.trim())
    .filter(Boolean);
}

function generateSchedules() {
  try {
    if (!responsesHeaders.length || !responsesRows.length) {
      throw new Error('Load a responses CSV first.');
    }
    if (!rehearsals.length) {
      throw new Error('Add at least one rehearsal first.');
    }

    const normalizedRehearsals = rehearsals.map(r => ({
      id: String(r.id || '').trim(),
      title: String(r.title || '').trim(),
      hours: Number(r.hours),
      people: splitPeopleList(r.neededPeople),
      mustFollowId: String(r.mustFollowId || '').trim(),
      mustFollowTitle: String(r.mustFollowTitle || '').trim()
    }));

    normalizedRehearsals.forEach(r => {
      if (!r.title) {
        throw new Error('Every rehearsal must have a title.');
      }
      if (!Number.isInteger(r.hours) || r.hours <= 0) {
        throw new Error(`Rehearsal "${r.title}" must have a whole-number hour length.`);
      }
      if (!r.people.length) {
        throw new Error(`Rehearsal "${r.title}" has no people selected.`);
      }
    });

    const daySpecs = getAvailabilityDaySpecsFromData();
    if (!daySpecs.length) {
      throw new Error('Could not find any availability columns in the CSV.');
    }

    const availabilityMap = buildAvailabilityMapFromData(daySpecs);
    const orderedRehearsals = topoSortRehearsals(normalizedRehearsals);
    const solutions = [];

    backtrackSchedules(orderedRehearsals, daySpecs, availabilityMap, {}, 0, solutions, MAX_SCHEDULE_SOLUTIONS);

    if (!solutions.length) {
      throw new Error('No valid schedules were found.');
    }

    const formattedSchedules = solutions.map((solution, index) => formatScheduleForDisplay(solution, daySpecs, index));
    generatedSchedules = dedupeSchedules(formattedSchedules);
    applyScheduleFilters();
    setStatus(`Generated ${generatedSchedules.length} schedule(s).`, 'success');
  } catch (error) {
    generatedSchedules = [];
    filteredScheduleIndexes = [];
    renderAll();
    setStatus(error.message || 'Failed to generate schedules.', 'error');
  }
}

function topoSortRehearsals(items) {
  const byId = {};
  const indegree = {};
  const graph = {};

  items.forEach(item => {
    byId[item.id] = item;
    indegree[item.id] = 0;
    graph[item.id] = [];
  });

  items.forEach(item => {
    if (item.mustFollowId && byId[item.mustFollowId]) {
      graph[item.mustFollowId].push(item.id);
      indegree[item.id] += 1;
    }
  });

  const queue = items.filter(item => indegree[item.id] === 0).map(item => item.id);
  const ordered = [];

  while (queue.length) {
    const id = queue.shift();
    ordered.push(byId[id]);

    graph[id].forEach(nextId => {
      indegree[nextId] -= 1;
      if (indegree[nextId] === 0) {
        queue.push(nextId);
      }
    });
  }

  if (ordered.length !== items.length) {
    throw new Error('Rehearsal dependencies contain a cycle.');
  }

  return ordered;
}

function backtrackSchedules(items, daySpecs, availabilityMap, placed, index, solutions, maxSolutions) {
  if (solutions.length >= maxSolutions) {
    return;
  }

  if (index >= items.length) {
    solutions.push(Object.values(placed).map(item => ({ ...item })));
    return;
  }

  const rehearsal = items[index];
  const candidates = getCandidateAssignments(rehearsal, daySpecs, availabilityMap, placed);

  for (const candidate of candidates) {
    placed[rehearsal.id] = candidate;
    backtrackSchedules(items, daySpecs, availabilityMap, placed, index + 1, solutions, maxSolutions);
    delete placed[rehearsal.id];

    if (solutions.length >= maxSolutions) {
      return;
    }
  }
}

function getCandidateAssignments(rehearsal, daySpecs, availabilityMap, placed) {
  const candidates = [];
  const prerequisite = rehearsal.mustFollowId ? placed[rehearsal.mustFollowId] : null;

  daySpecs.forEach(day => {
    const startLabels = getDayStartLabels(day.dayIndex, rehearsal.people, availabilityMap);

    startLabels.forEach(startLabel => {
      const startMinute = timeLabelToMinutes(startLabel);
      const endMinute = startMinute + (rehearsal.hours * 60);

      if (!allPeopleAvailableForBlock(rehearsal.people, day.dayIndex, startMinute, rehearsal.hours, availabilityMap)) {
        return;
      }

      const candidate = {
        rehearsalId: rehearsal.id,
        title: rehearsal.title,
        people: rehearsal.people.slice(),
        hours: rehearsal.hours,
        dayIndex: day.dayIndex,
        dayName: day.dayName,
        dateLabel: day.dateLabel,
        startMinute,
        endMinute,
        startLabel: minutesToTimeLabel(startMinute),
        endLabel: minutesToTimeLabel(endMinute)
      };

      if (prerequisite && !isAfterAssignment(candidate, prerequisite)) {
        return;
      }

      if (conflictsWithPlaced(candidate, placed)) {
        return;
      }

      candidates.push(candidate);
    });
  });

  candidates.sort((a, b) => {
    if (a.dayIndex !== b.dayIndex) {
      return a.dayIndex - b.dayIndex;
    }
    return a.startMinute - b.startMinute;
  });

  return candidates;
}

function getDayStartLabels(dayIndex, peopleList, availabilityMap) {
  const labels = new Set();

  peopleList.forEach(person => {
    const key = canonicalizeName(person);
    const personInfo = availabilityMap[key];
    if (!personInfo || !personInfo.byDay[dayIndex]) {
      return;
    }

    personInfo.byDay[dayIndex].forEach(label => labels.add(label));
  });

  return [...labels].sort((a, b) => timeLabelToMinutes(a) - timeLabelToMinutes(b));
}

function allPeopleAvailableForBlock(peopleList, dayIndex, startMinute, hours, availabilityMap) {
  for (const person of peopleList) {
    const key = canonicalizeName(person);
    const personInfo = availabilityMap[key];
    if (!personInfo || !personInfo.byDay[dayIndex]) {
      return false;
    }

    const available = personInfo.byDay[dayIndex];
    for (let i = 0; i < hours; i += 1) {
      const label = minutesToTimeLabel(startMinute + (i * 60));
      if (!available.has(normalizeTimeLabel(label))) {
        return false;
      }
    }
  }

  return true;
}

function isAfterAssignment(candidate, prerequisite) {
  if (candidate.dayIndex > prerequisite.dayIndex) {
    return true;
  }
  if (candidate.dayIndex < prerequisite.dayIndex) {
    return false;
  }
  return candidate.startMinute >= prerequisite.endMinute;
}

function conflictsWithPlaced(candidate, placed) {
  const placedAssignments = Object.values(placed);

  for (const other of placedAssignments) {
    if (candidate.dayIndex !== other.dayIndex) {
      continue;
    }

    const overlaps = candidate.startMinute < other.endMinute && other.startMinute < candidate.endMinute;
    if (overlaps) {
      return true;
    }
  }

  return false;
}

function formatScheduleForDisplay(solution, daySpecs, index) {
  const days = daySpecs.map(day => {
    const items = solution
      .filter(item => item.dayIndex === day.dayIndex)
      .sort((a, b) => a.startMinute - b.startMinute)
      .map(item => ({
        title: item.title,
        people: item.people.join(', '),
        start: item.startLabel,
        end: item.endLabel,
        hours: item.hours
      }));

    return {
      dayName: day.dayName,
      dateLabel: day.dateLabel,
      items
    };
  }).filter(day => day.items.length > 0);

  return {
    label: `Schedule ${index + 1}`,
    days
  };
}

function dedupeSchedules(schedules) {
  const seen = new Set();
  const result = [];

  schedules.forEach(schedule => {
    const key = JSON.stringify(
      schedule.days.map(day => ({
        dayName: day.dayName,
        dateLabel: day.dateLabel,
        items: day.items.map(item => ({
          title: item.title,
          people: item.people,
          start: item.start,
          end: item.end,
          hours: item.hours
        }))
      }))
    );

    if (!seen.has(key)) {
      seen.add(key);
      result.push({
        label: `Schedule ${result.length + 1}`,
        days: schedule.days
      });
    }
  });

  return result;
}

function populateSchedulePicker() {
  elements.schedulePicker.innerHTML = '';

  if (!filteredScheduleIndexes.length) {
    elements.schedulePicker.innerHTML = '<option value="">No matching schedules</option>';
    elements.scheduleCount.textContent = `0 of ${generatedSchedules.length} schedules shown`;
    return;
  }

  filteredScheduleIndexes.forEach(scheduleIndex => {
    const schedule = generatedSchedules[scheduleIndex];
    const option = document.createElement('option');
    option.value = String(scheduleIndex);
    option.textContent = schedule.label || `Schedule ${scheduleIndex + 1}`;
    elements.schedulePicker.appendChild(option);
  });

  elements.schedulePicker.value = String(filteredScheduleIndexes[0]);
  elements.scheduleCount.textContent = `${filteredScheduleIndexes.length} of ${generatedSchedules.length} schedules shown`;
}

function renderSelectedSchedule() {
  if (!generatedSchedules.length) {
    elements.scheduleWindow.innerHTML = '<div class="muted">No schedules generated yet.</div>';
    return;
  }

  const selectedIndex = Number(elements.schedulePicker.value);
  const schedule = generatedSchedules[selectedIndex];

  if (!schedule) {
    elements.scheduleWindow.innerHTML = '<div class="muted">No schedule selected.</div>';
    return;
  }

  const dayBlocks = schedule.days.map(day => {
    const rows = day.items.map(item => `
      <tr>
        <td>${escapeHtml(item.title)}</td>
        <td>${escapeHtml(item.people)}</td>
        <td>${escapeHtml(item.start)}</td>
        <td>${escapeHtml(item.end)}</td>
        <td>${escapeHtml(item.hours)}</td>
      </tr>
    `).join('');

    return `
      <div class="day-block">
        <div class="day-header">${escapeHtml(day.dayName)} (${escapeHtml(day.dateLabel)})</div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Title</th>
                <th>People Needed</th>
                <th>Start</th>
                <th>End</th>
                <th>Hours</th>
              </tr>
            </thead>
            <tbody>${rows}</tbody>
          </table>
        </div>
      </div>
    `;
  }).join('');

  elements.scheduleWindow.innerHTML = dayBlocks || '<div class="muted">This schedule has no items.</div>';
}

function applyScheduleFilters() {
  const selectedDay = elements.filterDay.value;
  const filterStart = elements.filterStart.value;
  const filterEnd = elements.filterEnd.value;

  filteredScheduleIndexes = generatedSchedules
    .map((schedule, index) => ({ schedule, index }))
    .filter(({ schedule }) => scheduleMatchesFilters(schedule, selectedDay, filterStart, filterEnd))
    .map(({ index }) => index);

  populateSchedulePicker();
  renderSelectedSchedule();
  updateFilterSummary();
}

function clearScheduleFilters() {
  elements.filterDay.value = '';
  elements.filterStart.value = '';
  elements.filterEnd.value = '';
  applyScheduleFilters();
}

function scheduleMatchesFilters(schedule, selectedDay, filterStart, filterEnd) {
  return schedule.days.some(day => {
    if (selectedDay && day.dayName !== selectedDay) {
      return false;
    }

    if (!filterStart && !filterEnd) {
      return true;
    }

    return day.items.some(item => itemMatchesTimeFilter(item, filterStart, filterEnd));
  });
}

function itemMatchesTimeFilter(item, filterStart, filterEnd) {
  const itemStartMinutes = hhmmToMinutes(labelToTimeInputValue(item.start));
  const itemEndMinutes = hhmmToMinutes(labelToTimeInputValue(item.end));

  if (filterStart && filterEnd) {
    const filterStartMinutes = hhmmToMinutes(filterStart);
    const filterEndMinutes = hhmmToMinutes(filterEnd);
    return itemStartMinutes < filterEndMinutes && filterStartMinutes < itemEndMinutes;
  }

  if (filterStart) {
    return itemEndMinutes > hhmmToMinutes(filterStart);
  }

  if (filterEnd) {
    return itemStartMinutes < hhmmToMinutes(filterEnd);
  }

  return true;
}

function labelToTimeInputValue(label) {
  const normalized = String(label).trim().toLowerCase();
  const match = normalized.match(/^(\d{1,2}):(\d{2})\s*([ap]m)$/);
  if (!match) {
    return '00:00';
  }

  let hour = Number(match[1]);
  const minute = match[2];
  const suffix = match[3];

  if (hour === 12) {
    hour = 0;
  }
  if (suffix === 'pm') {
    hour += 12;
  }

  return `${String(hour).padStart(2, '0')}:${minute}`;
}

function hhmmToMinutes(value) {
  const parts = String(value).split(':');
  if (parts.length !== 2) {
    return 0;
  }
  return Number(parts[0]) * 60 + Number(parts[1]);
}

function updateFilterSummary() {
  const parts = [];
  if (elements.filterDay.value) {
    parts.push(`day: ${elements.filterDay.value}`);
  }
  if (elements.filterStart.value || elements.filterEnd.value) {
    if (elements.filterStart.value && elements.filterEnd.value) {
      parts.push(`time: ${elements.filterStart.value}–${elements.filterEnd.value}`);
    } else if (elements.filterStart.value) {
      parts.push(`after ${elements.filterStart.value}`);
    } else {
      parts.push(`before ${elements.filterEnd.value}`);
    }
  }

  elements.filterSummary.textContent = parts.length
    ? `Filtered by ${parts.join(', ')}.`
    : 'Showing all generated schedules.';
}

function saveRehearsalsToStorage() {
  localStorage.setItem(STORAGE_KEY_REHEARSALS, JSON.stringify(rehearsals));
}

function loadRehearsalsFromStorage() {
  try {
    rehearsals = JSON.parse(localStorage.getItem(STORAGE_KEY_REHEARSALS) || '[]');
  } catch {
    rehearsals = [];
  }
  refreshMustFollowTitles();
}

function exportRehearsalsJson() {
  const blob = new Blob([JSON.stringify(rehearsals, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = 'rehearsals.json';
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function importRehearsalsJson(file) {
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const parsed = JSON.parse(String(reader.result || '[]'));
      if (!Array.isArray(parsed)) {
        throw new Error('The JSON file must contain an array of rehearsals.');
      }
      rehearsals = parsed;
      refreshMustFollowTitles();
      saveRehearsalsToStorage();
      resetGeneratedSchedules();
      resetForm();
      renderAll();
      setStatus('Imported rehearsals JSON.', 'success');
    } catch (error) {
      setStatus(error.message || 'Failed to import rehearsals JSON.', 'error');
    }
  };
  reader.readAsText(file);
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
