const FIREBASE_CONFIG = {
  apiKey: "AIzaSyCW3LHbfqyaSfoTi7lq0QGuApGKlry7y-I",
  authDomain: "cd-planning-tools.firebaseapp.com",
  projectId: "cd-planning-tools",
  storageBucket: "cd-planning-tools.firebasestorage.app",
  messagingSenderId: "997238317053",
  appId: "1:997238317053:web:502925ea322359cf63c394"
};
const FIRESTORE_COLLECTION = 'marketingCampaigns';
const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const COLORS = ['#004E78','#25A046','#00A7E1','#898989','#4A4A49','#C4923F'];
const STATUSES = ['Planned','In Progress','Approved','Completed','On Hold'];
const PRIORITIES = ['High','Medium','Low'];
const CAMPAIGN_TYPES = ['General','Product Launch','Seasonal','Trade Show','Email','Social','Print'];

const defaultConfig = {
  calendar_title: 'Marketing Calendar 2026',
  background_color: '#F5F5F5',
  surface_color: '#DFEEF6',
  text_color: '#4A4A49',
  primary_action_color: '#004E78',
  secondary_action_color: '#25A046',
  font_family: 'Proxima Nova',
  font_size: 14
};

let allTasks = [];
let editingTask = null;
let editColor = COLORS[0];
let dragId = null;
let deleteConfirmId = null;
let pendingHighlightId = null;
let authInstance = null;
let currentUser = null;
let firestoreDb = null;
let firestoreUnsubscribe = null;
let isFirestoreReady = false;
let isApplyingRemoteSnapshot = false;

function formatSavedTime(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return 'Not saved yet';
  return `Last saved at ${date.toLocaleString([], { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' })}`;
}

function formatChangeTime(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return '';
  return date.toLocaleString([], { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' });
}

function setLastSavedNote(value) {
  const note = document.getElementById('lastSavedNote');
  if (!note) return;
  note.textContent = value ? formatSavedTime(value) : 'Not saved yet';
}

function isSignedIn() {
  return Boolean(currentUser);
}

function requireSignedIn(actionText = 'edit shared campaigns') {
  if (isSignedIn()) return true;
  showToast(`Sign in with Google to ${actionText}`);
  return false;
}

function updateAuthUi() {
  const userLabel = document.getElementById('authUserLabel');
  const helperText = document.getElementById('authHelperText');
  const actionBtn = document.getElementById('authActionBtn');
  if (!userLabel || !helperText || !actionBtn) return;

  if (currentUser) {
    userLabel.textContent = currentUser.email || currentUser.displayName || 'Signed in';
    helperText.textContent = 'Shared Firestore sync is enabled';
    actionBtn.textContent = 'Sign Out';
    actionBtn.setAttribute('onclick', 'signOutUser()');
  } else {
    userLabel.textContent = 'Not signed in';
    helperText.textContent = 'Sign in to view and edit shared campaigns';
    actionBtn.textContent = 'Sign In With Google';
    actionBtn.setAttribute('onclick', 'signInWithGoogle()');
  }
}

function getCurrentActor() {
  return currentUser?.email || currentUser?.displayName || currentUser?.uid || 'Unknown';
}

function getLastChangeLabel(task) {
  const action = String(task.lastChangedAction || '').trim();
  const by = String(task.lastChangedBy || '').trim();
  const at = formatChangeTime(task.lastChangedAt);
  if (!action && !by && !at) return '-';
  return [action || 'Updated', by ? `by ${by}` : '', at ? `on ${at}` : ''].filter(Boolean).join(' ');
}

function generateId() {
  if (window.crypto && typeof window.crypto.randomUUID === 'function') {
    return window.crypto.randomUUID();
  }
  return `task-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function normalizeMonthValue(value) {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return Math.max(0, Math.min(MONTHS.length - 1, Math.trunc(value)));
  }

  const raw = String(value ?? '').trim();
  if (!raw) return 0;

  if (/^\d+$/.test(raw)) {
    return Math.max(0, Math.min(MONTHS.length - 1, parseInt(raw, 10)));
  }

  const index = MONTHS.findIndex(month => month.toLowerCase() === raw.toLowerCase());
  return index >= 0 ? index : 0;
}

function getCalendarYear() {
  const match = String(defaultConfig.calendar_title || '').match(/\b(20\d{2})\b/);
  return match ? parseInt(match[1], 10) : new Date().getFullYear();
}

function getDefaultDateForMonth(month) {
  const normalizedMonth = normalizeMonthValue(month);
  return formatDateInputValue(getCalendarYear(), normalizedMonth, 1);
}

function formatDateInputValue(year, month, day) {
  return `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
}

function createLocalDate(year, month, day) {
  const date = new Date(year, month, day);
  return date.getFullYear() === year && date.getMonth() === month && date.getDate() === day ? date : null;
}

function parseDateOnly(value) {
  const raw = String(value ?? '').trim();
  if (!raw) return null;

  const isoMatch = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:\b|T)/);
  if (isoMatch) {
    return createLocalDate(
      parseInt(isoMatch[1], 10),
      parseInt(isoMatch[2], 10) - 1,
      parseInt(isoMatch[3], 10)
    );
  }

  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) return null;
  return createLocalDate(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

function normalizeDateValue(value) {
  const raw = String(value ?? '').trim();
  const parsed = parseDateOnly(raw);
  return parsed ? formatDateInputValue(parsed.getFullYear(), parsed.getMonth(), parsed.getDate()) : raw;
}

function moveDateToMonth(value, month) {
  if (!value) return getDefaultDateForMonth(month);
  const parsed = parseDateOnly(value);
  if (!parsed) return getDefaultDateForMonth(month);

  const year = parsed.getFullYear();
  const targetMonth = normalizeMonthValue(month);
  const day = parsed.getDate();
  const maxDay = new Date(year, targetMonth + 1, 0).getDate();
  const safeDay = Math.min(day, maxDay);

  return formatDateInputValue(year, targetMonth, safeDay);
}

function getDateForSelectedMonth(value, month) {
  const targetMonth = normalizeMonthValue(month);
  if (!value) return getDefaultDateForMonth(targetMonth);

  const parsed = parseDateOnly(value);
  if (!parsed) return getDefaultDateForMonth(targetMonth);
  if (parsed.getMonth() === targetMonth) {
    return formatDateInputValue(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
  }

  return moveDateToMonth(value, targetMonth);
}

function normalizeStatus(value, fallback = '') {
  const raw = String(value ?? '').trim();
  if (!raw || raw.toLowerCase() === 'none') return '';
  const match = STATUSES.find(status => status.toLowerCase() === raw.toLowerCase());
  return match || fallback;
}

function normalizePriority(value, fallback = '') {
  const raw = String(value ?? '').trim();
  if (!raw || raw.toLowerCase() === 'none') return '';
  const match = PRIORITIES.find(priority => priority.toLowerCase() === raw.toLowerCase());
  return match || fallback;
}

function normalizeCampaignTypes(value, fallback = []) {
  const values = Array.isArray(value)
    ? value
    : String(value ?? '').split(/[;,|]/);
  if (!values.some(item => String(item).trim()) || values.some(item => String(item).trim().toLowerCase() === 'none')) {
    return fallback;
  }
  const selected = [];
  values.forEach(item => {
    const raw = String(item).trim();
    if (!raw) return;
    const builtIn = CAMPAIGN_TYPES.find(type => type.toLowerCase() === raw.toLowerCase());
    const valueToUse = builtIn || raw;
    if (!selected.some(existing => existing.toLowerCase() === valueToUse.toLowerCase())) {
      selected.push(valueToUse);
    }
  });
  return selected.length ? selected : fallback;
}

function getCampaignTypesLabel(value) {
  return normalizeCampaignTypes(value).join(', ');
}

function normalizeTask(task = {}, index = 0) {
  const name = String(task.name ?? task.text ?? 'Campaign').trim() || 'Campaign';
  const orderValue = Number(task.order);

  return {
    __backendId: String(task.__backendId ?? task.id ?? generateId()),
    month: normalizeMonthValue(task.month),
    order: Number.isFinite(orderValue) ? orderValue : Date.now() + index,
    deleted: Boolean(task.deleted),
    deletedAt: String(task.deletedAt ?? '').trim(),
    deletedBy: String(task.deletedBy ?? '').trim(),
    updatedAt: String(task.updatedAt ?? '').trim(),
    lastChangedAt: String(task.lastChangedAt ?? '').trim(),
    lastChangedBy: String(task.lastChangedBy ?? '').trim(),
    lastChangedAction: String(task.lastChangedAction ?? '').trim(),
    color: typeof task.color === 'string' && task.color ? task.color : COLORS[0],
    date: normalizeDateValue(task.date),
    status: Object.prototype.hasOwnProperty.call(task, 'status') ? normalizeStatus(task.status, STATUSES[0]) : STATUSES[0],
    priority: Object.prototype.hasOwnProperty.call(task, 'priority') ? normalizePriority(task.priority, 'Medium') : 'Medium',
    owner: String(task.owner ?? '').trim(),
    budget: String(task.budget ?? '').trim(),
    campaignTypes: Object.prototype.hasOwnProperty.call(task, 'campaignTypes') || Object.prototype.hasOwnProperty.call(task, 'campaignType') || Object.prototype.hasOwnProperty.call(task, 'type')
      ? normalizeCampaignTypes(task.campaignTypes ?? task.campaignType ?? task.type, [])
      : [CAMPAIGN_TYPES[0]],
    name,
    text: name,
    products: String(task.products ?? '').trim(),
    notes: String(task.notes ?? '').trim()
  };
}

function getFirestoreCollection() {
  if (!firestoreDb) return null;
  return firestoreDb.collection(FIRESTORE_COLLECTION);
}

function getFirestoreTaskPayload(task) {
  const normalized = normalizeTask(task);
  const { __backendId, ...payload } = normalized;
  return {
    ...payload,
    updatedAt: new Date().toISOString()
  };
}

async function saveTaskToFirestore(task) {
  const collection = getFirestoreCollection();
  if (!collection || !isFirestoreReady || isApplyingRemoteSnapshot) return;

  try {
    await collection.doc(task.__backendId).set(getFirestoreTaskPayload(task), { merge: true });
    setLastSavedNote(new Date().toISOString());
  } catch (error) {
    console.error('Failed to save campaign to Firestore:', error);
    showToast('Could not save to Firestore');
  }
}

async function deleteTaskFromFirestore(taskId) {
  const collection = getFirestoreCollection();
  if (!collection || !isFirestoreReady || isApplyingRemoteSnapshot) return;

  try {
    await collection.doc(taskId).delete();
  } catch (error) {
    console.error('Failed to delete campaign from Firestore:', error);
    showToast('Deleted locally only. Firestore unavailable.');
  }
}

async function syncAllTasksToFirestore() {
  const collection = getFirestoreCollection();
  if (!collection || !isFirestoreReady || isApplyingRemoteSnapshot) return;

  try {
    const snapshot = await collection.get();
    const batch = firestoreDb.batch();
    snapshot.docs.forEach(doc => batch.delete(doc.ref));
    allTasks.forEach(task => {
      batch.set(collection.doc(task.__backendId), getFirestoreTaskPayload(task), { merge: true });
    });
    await batch.commit();
    setLastSavedNote(new Date().toISOString());
  } catch (error) {
    console.error('Failed to sync campaigns to Firestore:', error);
    showToast('Firestore sync failed');
  }
}

function refreshTasks(tasks) {
  allTasks = tasks.map((task, index) => normalizeTask(task, index));
  renderTasks();
}

function getActiveTasks() {
  return allTasks.filter(task => !task.deleted);
}

function getDeletedTasks() {
  return allTasks
    .filter(task => task.deleted)
    .sort((a, b) => String(b.deletedAt || '').localeCompare(String(a.deletedAt || '')));
}

function upsertTask(updatedTask, actionOverride = '') {
  const exists = allTasks.some(task => task.__backendId === String(updatedTask.__backendId ?? updatedTask.id ?? ''));
  const timestamp = new Date().toISOString();
  const action = actionOverride || (exists ? 'Updated' : 'Created');
  const normalized = normalizeTask({
    ...updatedTask,
    updatedAt: timestamp,
    lastChangedAt: timestamp,
    lastChangedBy: getCurrentActor(),
    lastChangedAction: action
  });
  const nextTasks = exists
    ? allTasks.map(task => task.__backendId === normalized.__backendId ? normalized : task)
    : [...allTasks, normalized];

  allTasks = nextTasks;
  saveTaskToFirestore(normalized);
  renderTasks();
  return normalized;
}

function deleteTask(taskId) {
  if (!requireSignedIn('delete campaigns')) return false;
  const task = allTasks.find(item => item.__backendId === taskId);
  if (!task) return false;
  upsertTask({
    ...task,
    deleted: true,
    deletedAt: new Date().toISOString(),
    deletedBy: currentUser?.email || currentUser?.uid || ''
  }, 'Deleted');
  return true;
}

function restoreTask(taskId) {
  if (!requireSignedIn('restore campaigns')) return false;
  const task = allTasks.find(item => item.__backendId === taskId);
  if (!task) return false;
  upsertTask({
    ...task,
    deleted: false,
    deletedAt: '',
    deletedBy: ''
  }, 'Restored');
  showToast('Campaign restored');
  return true;
}

function nextOrder(month) {
  const monthTasks = getActiveTasks().filter(task => task.month === month);
  if (!monthTasks.length) return Date.now();
  return Math.max(...monthTasks.map(task => Number(task.order) || 0)) + 1;
}

function initEditMonthOptions() {
  const select = document.getElementById('editMonth');
  select.innerHTML = MONTHS.map((month, index) => `<option value="${index}">${month}</option>`).join('');
}

function initEditDateControls() {
  const monthSelect = document.getElementById('editMonth');
  const dateInput = document.getElementById('editDate');
  monthSelect.addEventListener('change', () => {
    dateInput.value = getDateForSelectedMonth(dateInput.value, monthSelect.value);
  });
  dateInput.addEventListener('change', () => {
    const parsed = parseDateOnly(dateInput.value);
    if (parsed) {
      monthSelect.value = String(parsed.getMonth());
    }
  });
}

function getCampaignTypeOptions(extraTypes = []) {
  const usedTypes = getActiveTasks().flatMap(task => normalizeCampaignTypes(task.campaignTypes ?? task.campaignType, []));
  const options = [];
  [...CAMPAIGN_TYPES, ...usedTypes, ...extraTypes].forEach(type => {
    const value = String(type || '').trim();
    if (value && !options.some(option => option.toLowerCase() === value.toLowerCase())) {
      options.push(value);
    }
  });
  return options;
}

function renderCampaignTypeCheckboxes(selectedTypes = []) {
  const container = document.getElementById('editCampaignTypes');
  if (!container) return;

  const options = ['', ...getCampaignTypeOptions(selectedTypes)];
  container.innerHTML = options.map(type => {
    const label = type || 'None';
    return `
      <label class="checkbox-option">
        <input type="checkbox" value="${escHtml(type)}">
        <span>${escHtml(label)}</span>
      </label>
    `;
  }).join('');

  container.querySelectorAll('input[type="checkbox"]').forEach(checkbox => {
    checkbox.addEventListener('change', () => {
      const noneCheckbox = container.querySelector('input[value=""]');
      const typeCheckboxes = Array.from(container.querySelectorAll('input[type="checkbox"]')).filter(input => input.value);

      if (!checkbox.value && checkbox.checked) {
        typeCheckboxes.forEach(input => { input.checked = false; });
        return;
      }

      if (checkbox.value && checkbox.checked && noneCheckbox) {
        noneCheckbox.checked = false;
      }

      if (noneCheckbox && !typeCheckboxes.some(input => input.checked)) {
        noneCheckbox.checked = true;
      }
    });
  });

  container.querySelectorAll('input[type="checkbox"]').forEach(checkbox => {
    checkbox.checked = selectedTypes.length ? selectedTypes.includes(checkbox.value) : checkbox.value === '';
  });
}

function initCampaignTypeCheckboxes() {
  renderCampaignTypeCheckboxes();
}

function addCustomCampaignType() {
  const input = document.getElementById('editCustomCampaignType');
  if (!input) return;
  const customType = input.value.trim();
  if (!customType) return;

  const selectedTypes = Array.from(document.querySelectorAll('#editCampaignTypes input[type="checkbox"]:checked'))
    .map(checkbox => checkbox.value)
    .filter(Boolean);
  const nextTypes = normalizeCampaignTypes([...selectedTypes, customType], []);
  renderCampaignTypeCheckboxes(nextTypes);
  input.value = '';
}

function setSelectOptions(selectId, values, emptyLabel) {
  const select = document.getElementById(selectId);
  if (!select) return;
  const currentValue = select.value;
  select.innerHTML = `<option value="">${emptyLabel}</option>` + values
    .map(value => `<option value="${escHtml(value)}">${escHtml(value)}</option>`)
    .join('');
  select.value = values.includes(currentValue) ? currentValue : '';
}

function initFilterOptions() {
  setSelectOptions('statusFilter', STATUSES, 'All statuses');
  setSelectOptions('priorityFilter', PRIORITIES, 'All priorities');
  updateOwnerFilterOptions();
  updateTypeFilterOptions();

  ['searchInput', 'statusFilter', 'priorityFilter', 'ownerFilter', 'typeFilter'].forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('input', renderTasks);
      el.addEventListener('change', renderTasks);
    }
  });

  document.getElementById('clearFiltersBtn').addEventListener('click', () => {
    document.getElementById('searchInput').value = '';
    document.getElementById('statusFilter').value = '';
    document.getElementById('priorityFilter').value = '';
    document.getElementById('ownerFilter').value = '';
    document.getElementById('typeFilter').value = '';
    renderTasks();
  });
}

function initFilterToggle() {
  const panel = document.querySelector('.filter-panel');
  const toggle = document.getElementById('filterToggle');
  if (!panel || !toggle) return;

  toggle.addEventListener('click', () => {
    const isCollapsed = panel.classList.toggle('is-collapsed');
    toggle.textContent = isCollapsed ? 'Show' : 'Hide';
    toggle.setAttribute('aria-expanded', String(!isCollapsed));
  });
}

function initUpcomingToggle() {
  const panel = document.querySelector('.upcoming-panel');
  const toggle = document.getElementById('upcomingToggle');
  if (!panel || !toggle) return;

  toggle.addEventListener('click', () => {
    const isCollapsed = panel.classList.toggle('is-collapsed');
    toggle.textContent = isCollapsed ? 'Show' : 'Hide';
    toggle.setAttribute('aria-expanded', String(!isCollapsed));
  });
}

function initQuarterToggle() {
  const panel = document.querySelector('.quarter-panel');
  const toggle = document.getElementById('quarterToggle');
  if (!panel || !toggle) return;

  toggle.addEventListener('click', () => {
    const isCollapsed = panel.classList.toggle('is-collapsed');
    toggle.textContent = isCollapsed ? 'Show' : 'Hide';
    toggle.setAttribute('aria-expanded', String(!isCollapsed));
  });
}

function initDeletedToggle() {
  const panel = document.querySelector('.deleted-panel');
  const toggle = document.getElementById('deletedToggle');
  if (!panel || !toggle) return;

  toggle.addEventListener('click', () => {
    const isCollapsed = panel.classList.toggle('is-collapsed');
    toggle.textContent = isCollapsed ? 'Show' : 'Hide';
    toggle.setAttribute('aria-expanded', String(!isCollapsed));
  });
}

function updateOwnerFilterOptions() {
  const owners = [...new Set(getActiveTasks().map(task => String(task.owner || '').trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b));
  setSelectOptions('ownerFilter', owners, 'All owners');
}

function updateTypeFilterOptions() {
  setSelectOptions('typeFilter', getCampaignTypeOptions(), 'All types');
}

function getFilterState() {
  return {
    query: String(document.getElementById('searchInput')?.value || '').trim().toLowerCase(),
    status: document.getElementById('statusFilter')?.value || '',
    priority: document.getElementById('priorityFilter')?.value || '',
    owner: document.getElementById('ownerFilter')?.value || '',
    type: document.getElementById('typeFilter')?.value || ''
  };
}

function hasActiveFilters(filters = getFilterState()) {
  return Boolean(filters.query || filters.status || filters.priority || filters.owner || filters.type);
}

function taskMatchesFilters(task, filters) {
  const campaignTypes = normalizeCampaignTypes(task.campaignTypes ?? task.campaignType);
  const searchable = [
    task.name,
    task.text,
    task.date,
    task.owner,
    task.budget,
    task.products,
    task.notes,
    normalizeStatus(task.status),
    normalizePriority(task.priority),
    campaignTypes.join(' '),
    MONTHS[normalizeMonthValue(task.month)]
  ].join(' ').toLowerCase();

  if (filters.query && !searchable.includes(filters.query)) return false;
  if (filters.status && normalizeStatus(task.status) !== filters.status) return false;
  if (filters.priority && normalizePriority(task.priority) !== filters.priority) return false;
  if (filters.owner && String(task.owner || '').trim() !== filters.owner) return false;
  if (filters.type && !campaignTypes.includes(filters.type)) return false;
  return true;
}

function getVisibleTasks(filters = getFilterState()) {
  return getActiveTasks().filter(task => taskMatchesFilters(task, filters));
}

function getUpcomingTasks(tasks, limit = 5) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  return tasks
    .map(task => ({ task, date: parseCampaignDate(task.date) }))
    .filter(item => item.date && item.date >= today)
    .sort((a, b) => a.date - b.date || Number(a.task.order || 0) - Number(b.task.order || 0))
    .slice(0, limit);
}

function renderUpcomingDeadlines(tasks, filtersActive) {
  const list = document.getElementById('upcomingList');
  const subtitle = document.getElementById('upcomingSubtitle');
  if (!list || !subtitle) return;

  const upcoming = getUpcomingTasks(tasks);
  subtitle.textContent = filtersActive ? 'Next 5 matching dated campaigns' : 'Next 5 dated campaigns';
  list.innerHTML = '';

  if (!upcoming.length) {
    const empty = document.createElement('div');
    empty.className = 'upcoming-empty';
    empty.textContent = filtersActive ? 'No upcoming deadlines match the current filters.' : 'No upcoming dated campaigns yet.';
    list.appendChild(empty);
    return;
  }

  upcoming.forEach(({ task, date }) => {
    const item = document.createElement('button');
    const meta = [
      MONTHS[normalizeMonthValue(task.month)],
      normalizeStatus(task.status),
      normalizePriority(task.priority)
    ].filter(Boolean).join(' Â· ');
    item.type = 'button';
    item.className = 'upcoming-item';
    item.style.borderLeft = `4px solid ${task.color || COLORS[0]}`;
    item.innerHTML = `
      <div class="upcoming-date">${escHtml(formatCardDate(task.date))} Â· ${escHtml(formatDeadlineDistance(date))}</div>
      <div class="upcoming-name">${escHtml(task.name || task.text || 'Campaign')}</div>
      <div class="upcoming-meta">${escHtml(meta)}</div>
    `;
    item.addEventListener('click', () => openView(task));
    list.appendChild(item);
  });
}

function getQuarterStats(tasks) {
  return [0, 1, 2, 3].map(quarter => {
    const startMonth = quarter * 3;
    const quarterTasks = tasks.filter(task => {
      const month = normalizeMonthValue(task.month);
      return month >= startMonth && month <= startMonth + 2;
    });
    const productCount = quarterTasks.reduce((total, task) => total + getProductsList(task.products).length, 0);
    const highPriorityCount = quarterTasks.filter(task => normalizePriority(task.priority) === 'High').length;

    return {
      label: `Q${quarter + 1}`,
      months: MONTHS.slice(startMonth, startMonth + 3).map(month => month.slice(0, 3)).join(' - '),
      campaignCount: quarterTasks.length,
      productCount,
      highPriorityCount
    };
  });
}

function renderQuarterView(tasks, filtersActive) {
  const grid = document.getElementById('quarterGrid');
  const subtitle = document.getElementById('quarterSubtitle');
  if (!grid || !subtitle) return;

  subtitle.textContent = filtersActive ? 'Filtered campaigns grouped by quarter' : 'Campaigns grouped by quarter';
  grid.innerHTML = getQuarterStats(tasks).map(quarter => `
    <div class="quarter-card">
      <div class="quarter-name">${escHtml(quarter.label)}</div>
      <div class="quarter-months">${escHtml(quarter.months)}</div>
      <div class="quarter-stats">
        <span class="quarter-stat">${quarter.campaignCount} campaign${quarter.campaignCount === 1 ? '' : 's'}</span>
        <span class="quarter-stat">${quarter.productCount} product${quarter.productCount === 1 ? '' : 's'}</span>
        <span class="quarter-stat">${quarter.highPriorityCount} high priority</span>
      </div>
    </div>
  `).join('');
}

function renderDeletedPanel() {
  const list = document.getElementById('deletedList');
  const subtitle = document.getElementById('deletedSubtitle');
  if (!list || !subtitle) return;

  if (!isSignedIn()) {
    subtitle.textContent = 'Sign in to view deleted campaigns';
    list.innerHTML = '';
    const empty = document.createElement('div');
    empty.className = 'deleted-empty';
    empty.textContent = 'No deleted campaigns visible while signed out.';
    list.appendChild(empty);
    return;
  }

  const deletedTasks = getDeletedTasks();
  subtitle.textContent = deletedTasks.length
    ? `${deletedTasks.length} deleted campaign${deletedTasks.length === 1 ? '' : 's'} available to restore`
    : 'Restore deleted campaigns';

  list.innerHTML = '';
  if (!deletedTasks.length) {
    const empty = document.createElement('div');
    empty.className = 'deleted-empty';
    empty.textContent = 'No deleted campaigns.';
    list.appendChild(empty);
    return;
  }

  deletedTasks.slice(0, 10).forEach(task => {
    const item = document.createElement('div');
    item.className = 'deleted-item';
    item.innerHTML = `
      <div>
        <div class="deleted-name">${escHtml(task.name || task.text || 'Campaign')}</div>
        <div class="deleted-meta">
          ${escHtml([
            task.deletedAt ? `Deleted ${formatSavedTime(task.deletedAt).replace('Last saved at ', '')}` : '',
            task.deletedBy ? `by ${task.deletedBy}` : ''
          ].filter(Boolean).join(' '))}
        </div>
      </div>
      <button class="restore-btn" type="button">Restore</button>
    `;
    item.querySelector('.restore-btn').addEventListener('click', () => restoreTask(task.__backendId));
    list.appendChild(item);
  });
}

function flashSavedTask(taskId) {
  pendingHighlightId = taskId;
  renderTasks();
  setTimeout(() => {
    if (pendingHighlightId === taskId) {
      pendingHighlightId = null;
      const card = document.querySelector(`.task-card[data-id="${taskId}"]`);
      if (card) card.classList.remove('just-saved');
    }
  }, 1700);
}

function getColorLabel(color) {
  return color === '#004E78' ? 'Regal Blue'
    : color === '#25A046' ? 'Classic Green'
    : color === '#00A7E1' ? 'Bright Blue'
    : color === '#898989' ? 'Grey'
    : color === '#4A4A49' ? 'Dark Grey'
    : 'Holiday Gold';
}

function formatCardDate(value) {
  if (!value) return '';
  const date = parseDateOnly(value);
  if (!date) return String(value);
  return date.toLocaleDateString([], { month: 'short', day: 'numeric' });
}

function parseCampaignDate(value) {
  if (!value) return null;
  return parseDateOnly(value);
}

function getDaysUntil(date) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const target = new Date(date);
  target.setHours(0, 0, 0, 0);
  return Math.round((target - today) / 86400000);
}

function formatDeadlineDistance(date) {
  const days = getDaysUntil(date);
  if (days === 0) return 'Today';
  if (days === 1) return 'Tomorrow';
  if (days > 1) return `In ${days} days`;
  if (days === -1) return 'Yesterday';
  return `${Math.abs(days)} days ago`;
}

function getProductsList(value) {
  return String(value || '')
    .split(/\r?\n|,/)
    .map(item => item.trim())
    .filter(Boolean);
}

function getProductSummary(value) {
  const items = getProductsList(value);
  if (!items.length) return '';
  if (items.length === 1) return items[0];
  return `${items.length} products`;
}

function getMonthSummary(tasks) {
  const campaignText = `${tasks.length} campaign${tasks.length === 1 ? '' : 's'}`;
  const productCount = tasks.reduce((total, task) => total + getProductsList(task.products).length, 0);
  if (!productCount) return campaignText;
  return `${campaignText} Â· ${productCount} product${productCount === 1 ? '' : 's'}`;
}

function getNotesPreview(value) {
  return String(value || '').replace(/\s+/g, ' ').trim();
}

function getStatusClass(status) {
  return `status-${normalizeStatus(status).toLowerCase().replace(/\s+/g, '-')}`;
}

function getPriorityClass(priority) {
  return `priority-${normalizePriority(priority).toLowerCase()}`;
}

function renderCardMarkup(t) {
  const displayDate = formatCardDate(t.date);
  const productSummary = getProductSummary(t.products);
  const notesPreview = getNotesPreview(t.notes);
  const status = normalizeStatus(t.status);
  const priority = normalizePriority(t.priority);
  const owner = String(t.owner || '').trim();
  const budget = String(t.budget || '').trim();
  const campaignTypesLabel = getCampaignTypesLabel(t.campaignTypes ?? t.campaignType);

  return `
    <div class="task-content">
      <div class="task-title">${escHtml(t.text)}</div>
      ${displayDate ? `<div class="task-date">${escHtml(displayDate)}</div>` : ''}
      <div class="task-meta">
        ${status ? `<span class="task-pill status-pill ${getStatusClass(status)}">${escHtml(status)}</span>` : ''}
        ${priority ? `<span class="task-pill priority-pill ${getPriorityClass(priority)}">${escHtml(priority)}</span>` : ''}
        ${owner ? `<span class="task-pill">Owner: ${escHtml(owner)}</span>` : ''}
        ${budget ? `<span class="task-pill">Budget: ${escHtml(budget)}</span>` : ''}
        ${campaignTypesLabel ? `<span class="task-pill">${escHtml(campaignTypesLabel)}</span>` : ''}
        ${productSummary ? `<span class="task-pill">${escHtml(productSummary)}</span>` : ''}
      </div>
      ${notesPreview ? `<div class="task-notes-preview">${escHtml(notesPreview)}</div>` : ''}
    </div>
    <div class="actions" style="display:flex;gap:4px;flex-shrink:0;">
      <button class="view-btn" style="background:none;border:none;cursor:pointer;padding:2px;color:#004E78;"><svg data-lucide="eye" style="width:14px;height:14px;" xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg></button>
      <button class="edit-btn" style="background:none;border:none;cursor:pointer;padding:2px;color:#004E78;"><svg data-lucide="pencil" style="width:14px;height:14px;" xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 3a2.828 2.828 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5L17 3z"></path></svg></button>
      <button class="del-btn" style="background:none;border:none;cursor:pointer;padding:2px;color:#898989;"><svg data-lucide="trash-2" style="width:14px;height:14px;" xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path><line x1="10" y1="11" x2="10" y2="17"></line><line x1="14" y1="11" x2="14" y2="17"></line></svg></button>
    </div>
  `;
}

// Build month grid
function buildGrid() {
  const grid = document.getElementById('monthGrid');
  grid.innerHTML = '';
  MONTHS.forEach((m, i) => {
    const col = document.createElement('div');
    col.className = 'month-col';
    col.dataset.month = i;
    col.style.background = defaultConfig.surface_color;
    col.innerHTML = `
      <div class="month-title">${m}</div>
      <div class="month-summary is-empty" data-summary-month="${i}">0 campaigns</div>
      <div class="task-list" data-month="${i}"></div>
      <button class="add-btn" data-month="${i}" style="color:${defaultConfig.secondary_action_color};border-color:${defaultConfig.secondary_action_color}33;">+ Add Campaign</button>
    `;
    // Drag events
    col.addEventListener('dragover', e => { e.preventDefault(); col.classList.add('drag-over'); });
    col.addEventListener('dragleave', () => col.classList.remove('drag-over'));
    col.addEventListener('drop', async e => {
      e.preventDefault(); 
      col.classList.remove('drag-over');
      if (!dragId) return;
      if (!requireSignedIn('reorder campaigns')) { dragId = null; return; }
      const task = allTasks.find(t => t.__backendId === dragId);
      if (task) {
        upsertTask({
          ...task,
          month: i,
          date: moveDateToMonth(task.date, i),
          order: nextOrder(i)
        }, 'Moved');
      }
      dragId = null;
    });
    // Add button
    col.querySelector('.add-btn').addEventListener('click', () => addTask(i));
    grid.appendChild(col);
  });
}

async function addTask(month) {
  if (!requireSignedIn('add campaigns')) return;
  if (allTasks.length >= 999) { showToast('Task limit reached (999)'); return; }
  const task = {
    __backendId: generateId(),
    month: normalizeMonthValue(month),
    order: nextOrder(month),
    color: COLORS[Math.floor(Math.random() * COLORS.length)],
    date: '',
    status: STATUSES[0],
    priority: 'Medium',
    owner: '',
    budget: '',
    campaignTypes: [CAMPAIGN_TYPES[0]],
    name: '',
    text: '',
    products: '',
    notes: ''
  };
  openEdit(task);
}

function renderTasks() {
  updateOwnerFilterOptions();
  updateTypeFilterOptions();
  const filters = getFilterState();
  const visibleTasks = getVisibleTasks(filters);
  const filtersActive = hasActiveFilters(filters);
  const activeTasks = getActiveTasks();
  renderUpcomingDeadlines(visibleTasks, filtersActive);
  renderQuarterView(visibleTasks, filtersActive);
  renderDeletedPanel();

  document.querySelectorAll('.month-col').forEach(col => {
    const month = parseInt(col.dataset.month, 10);
    const list = col.querySelector('.task-list');
    const summary = col.querySelector('.month-summary');
    const tasks = visibleTasks.filter(t => normalizeMonthValue(t.month) === month).sort((a,b) => a.order - b.order);

    col.classList.toggle('has-tasks', tasks.length > 0);
    summary.textContent = getMonthSummary(tasks);
    summary.classList.toggle('is-empty', tasks.length === 0);
    list.innerHTML = '';

    if (!tasks.length) {
      const hint = document.createElement('div');
      hint.className = 'task-empty-hint';
      hint.textContent = filtersActive ? 'No matching campaigns' : 'No campaigns yet';
      list.appendChild(hint);
    }

    tasks.forEach(t => {
      const card = createCard(t);
      if (pendingHighlightId === t.__backendId) card.classList.add('just-saved');
      list.appendChild(card);
    });
  });
  document.getElementById('taskCount').textContent = filtersActive
    ? `${visibleTasks.length} of ${activeTasks.length} campaign${activeTasks.length !== 1 ? 's' : ''} shown`
    : `${activeTasks.length} campaign${activeTasks.length!==1?'s':''}`;
}

function createCard(t) {
  const card = document.createElement('div');
  card.className = 'task-card';
  card.dataset.id = t.__backendId;
  card.draggable = true;
  card.style.background = '#FFFFFF';
  card.style.borderLeft = `4px solid ${t.color || COLORS[0]}`;
  card.style.color = defaultConfig.text_color;
  card.innerHTML = renderCardMarkup(t);

  card.addEventListener('dragstart', e => { dragId = t.__backendId; card.classList.add('dragging'); e.dataTransfer.effectAllowed = 'move'; });
  card.addEventListener('dragend', () => { card.classList.remove('dragging'); dragId = null; });
  card.querySelector('.edit-btn').addEventListener('click', e => { e.stopPropagation(); openEdit(t); });
  card.querySelector('.view-btn').addEventListener('click', e => { e.stopPropagation(); openView(t); });
  card.querySelector('.del-btn').addEventListener('click', e => {
    e.stopPropagation();
    if (deleteConfirmId === t.__backendId) {
      // already confirming, do nothing
      return;
    }
    showDeleteConfirm(card, t);
  });
  card.addEventListener('click', () => openEdit(t));
  return card;
}

function showDeleteConfirm(card, t) {
  // Remove any previous confirm bars
  document.querySelectorAll('.confirm-bar').forEach(b => b.remove());
  deleteConfirmId = t.__backendId;
  const bar = document.createElement('div');
  bar.className = 'confirm-bar';
  bar.innerHTML = `
    <button style="background:#ef4444;color:#fff;">Delete</button>
    <button style="background:#3a3f4e;color:#ccc;">Cancel</button>
  `;
  bar.children[0].addEventListener('click', async e => {
    e.stopPropagation();
    bar.children[0].disabled = true; bar.children[0].textContent = '...';
    const deleted = deleteTask(t.__backendId);
    if (deleted) showToast('Campaign moved to Recently Deleted');
    deleteConfirmId = null;
  });
  bar.children[1].addEventListener('click', e => { e.stopPropagation(); bar.remove(); deleteConfirmId = null; });
  card.appendChild(bar);
}

function openEdit(t) {
  if (!requireSignedIn('edit campaigns')) return;
  editingTask = t;
  editColor = t.color || COLORS[0];
  
  // Ensure all values are strings and not undefined
  document.getElementById('editMonth').value = String(normalizeMonthValue(t.month));
  document.getElementById('editDate').value = getDateForSelectedMonth(t.date, t.month);
  document.getElementById('editName').value = t.name ? String(t.name) : '';
  document.getElementById('editStatus').value = normalizeStatus(t.status);
  document.getElementById('editPriority').value = normalizePriority(t.priority);
  document.getElementById('editOwner').value = t.owner ? String(t.owner) : '';
  document.getElementById('editBudget').value = t.budget ? String(t.budget) : '';
  const selectedCampaignTypes = normalizeCampaignTypes(t.campaignTypes ?? t.campaignType);
  renderCampaignTypeCheckboxes(selectedCampaignTypes);
  document.getElementById('editCustomCampaignType').value = '';
  document.getElementById('editProducts').value = t.products ? String(t.products) : '';
  document.getElementById('editNotes').value = t.notes ? String(t.notes) : '';
  
  renderEditColors();
  document.getElementById('editOverlay').classList.add('open');
  
  // Focus the name field so new campaigns are faster to add.
  setTimeout(() => {
    document.getElementById('editName').focus();
    document.getElementById('editName').select();
  }, 50);
}

function renderEditColors() {
  const container = document.getElementById('editColors');
  container.innerHTML = '';
  COLORS.forEach(c => {
    const dot = document.createElement('div');
    dot.className = 'color-dot' + (c===editColor?' active':'');
    dot.style.background = c;
    dot.addEventListener('click', () => { editColor = c; renderEditColors(); });
    container.appendChild(dot);
  });
}

function closeEdit() {
  document.getElementById('editOverlay').classList.remove('open');
  editingTask = null;
}

function openView(t) {
  document.getElementById('viewDate').textContent = t.date && String(t.date).trim() ? String(t.date) : '-';
  document.getElementById('viewName').textContent = t.name && String(t.name).trim() ? String(t.name) : '-';
  const status = normalizeStatus(t.status);
  const priority = normalizePriority(t.priority);
  const campaignTypesLabel = getCampaignTypesLabel(t.campaignTypes ?? t.campaignType);
  document.getElementById('viewStatus').textContent = status || '-';
  document.getElementById('viewPriority').textContent = priority || '-';
  document.getElementById('viewOwner').textContent = t.owner && String(t.owner).trim() ? String(t.owner) : '-';
  document.getElementById('viewBudget').textContent = t.budget && String(t.budget).trim() ? String(t.budget) : '-';
  document.getElementById('viewCampaignType').textContent = campaignTypesLabel || '-';
  document.getElementById('viewProducts').textContent = t.products && String(t.products).trim() ? String(t.products) : '-';
  document.getElementById('viewNotes').textContent = t.notes && String(t.notes).trim() ? String(t.notes) : '-';
  document.getElementById('viewLastChange').textContent = getLastChangeLabel(t);
  document.getElementById('viewColor').style.background = t.color || COLORS[0];
  document.getElementById('viewOverlay').classList.add('open');
}

function closeView() {
  document.getElementById('viewOverlay').classList.remove('open');
}

async function saveEdit() {
  if (!editingTask) return;
  if (!requireSignedIn('save campaigns')) return;
  const month = parseInt(document.getElementById('editMonth').value, 10) || 0;
  const date = document.getElementById('editDate').value || '';
  const name = document.getElementById('editName').value.trim() || 'Untitled';
  const status = normalizeStatus(document.getElementById('editStatus').value);
  const priority = normalizePriority(document.getElementById('editPriority').value);
  const owner = document.getElementById('editOwner').value.trim() || '';
  const budget = document.getElementById('editBudget').value.trim() || '';
  const campaignTypes = normalizeCampaignTypes(Array.from(document.querySelectorAll('#editCampaignTypes input[type="checkbox"]:checked')).map(input => input.value));
  const products = document.getElementById('editProducts').value.trim() || '';
  const notes = document.getElementById('editNotes').value.trim() || '';
  const btn = document.querySelector('.edit-modal button:last-child');
  btn.disabled = true; btn.textContent = '...';
  btn.disabled = false; btn.textContent = 'Save';
  editingTask = upsertTask({ ...editingTask, month, text: name, date, status, priority, owner, budget, campaignTypes, name, products, notes, color: editColor });
  closeEdit();
  flashSavedTask(editingTask.__backendId);
  showToast('Campaign saved!');
}

function showToast(msg) {
  const el = document.getElementById('toast');
  el.textContent = msg; el.classList.add('show');
  setTimeout(() => el.classList.remove('show'), 2500);
}

function escHtml(s) { const d=document.createElement('div'); d.textContent=s; return d.innerHTML; }

function buildCSV() {
  let csvContent = 'Month,Campaign Name,Date,Status,Priority,Owner,Budget,Campaign Types,Products,Notes,Last Change,Last Changed By,Last Changed At,Category\n';
  
  MONTHS.forEach((month, idx) => {
    const tasks = getActiveTasks().filter(t => t.month === idx).sort((a,b) => a.order - b.order);
    tasks.forEach(t => {
      const name = t.name || '';
      const date = t.date || '';
      const status = normalizeStatus(t.status);
      const priority = normalizePriority(t.priority);
      const owner = t.owner || '';
      const budget = t.budget || '';
      const campaignTypes = normalizeCampaignTypes(t.campaignTypes ?? t.campaignType).join('; ');
      const products = t.products || '';
      const notes = t.notes || '';
      const lastChangedAction = t.lastChangedAction || '';
      const lastChangedBy = t.lastChangedBy || '';
      const lastChangedAt = t.lastChangedAt || '';
      const colorName = getColorLabel(t.color || COLORS[0]);
      
      csvContent += `"${csvEscape(month)}","${csvEscape(name)}","${csvEscape(date)}","${csvEscape(status)}","${csvEscape(priority)}","${csvEscape(owner)}","${csvEscape(budget)}","${csvEscape(campaignTypes)}","${csvEscape(products)}","${csvEscape(notes)}","${csvEscape(lastChangedAction)}","${csvEscape(lastChangedBy)}","${csvEscape(lastChangedAt)}","${csvEscape(colorName)}"\n`;
    });
  });
  return csvContent;
}

function downloadCalendar() {
  const csvContent = '\uFEFF' + buildCSV();
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `marketing-calendar-${new Date().toISOString().split('T')[0]}.csv`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
  showToast('Excel-friendly CSV exported!');
}

function copyCalendar() {
  const csvContent = buildCSV();
  navigator.clipboard.writeText(csvContent).then(() => {
    showToast('Calendar copied to clipboard!');
  }).catch(() => {
    showToast('Failed to copy');
  });
}

function csvEscape(str) {
  return String(str).replace(/"/g, '""');
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let value = '';
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];

    if (char === '"') {
      if (inQuotes && next === '"') {
        value += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (char === ',' && !inQuotes) {
      row.push(value);
      value = '';
      continue;
    }

    if ((char === '\n' || char === '\r') && !inQuotes) {
      if (char === '\r' && next === '\n') i += 1;
      row.push(value);
      if (row.some(cell => String(cell).trim() !== '')) rows.push(row);
      row = [];
      value = '';
      continue;
    }

    value += char;
  }

  row.push(value);
  if (row.some(cell => String(cell).trim() !== '')) rows.push(row);
  return rows;
}

function getHeaderIndex(headers, names) {
  return headers.findIndex(header => names.includes(String(header).trim().toLowerCase()));
}

function getCsvValue(row, headerIndex, fallbackIndex) {
  const index = headerIndex >= 0 ? headerIndex : fallbackIndex;
  return row[index] || '';
}

function importCalendar(event) {
  if (!requireSignedIn('import campaigns')) {
    event.target.value = '';
    return;
  }
  const file = event.target.files && event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = () => {
    try {
      const rows = parseCsv(String(reader.result || ''));
      if (!rows.length) throw new Error('The CSV file is empty.');

      const headers = rows[0].map(header => String(header).trim().toLowerCase());
      const monthIndex = getHeaderIndex(headers, ['month']);
      const nameIndex = getHeaderIndex(headers, ['campaign name', 'name', 'campaign']);
      const dateIndex = getHeaderIndex(headers, ['date']);
      const statusIndex = getHeaderIndex(headers, ['status', 'campaign status']);
      const priorityIndex = getHeaderIndex(headers, ['priority', 'campaign priority']);
      const ownerIndex = getHeaderIndex(headers, ['owner', 'responsible person', 'responsible', 'assigned to']);
      const budgetIndex = getHeaderIndex(headers, ['budget', 'cost', 'estimated budget']);
      const campaignTypesIndex = getHeaderIndex(headers, ['campaign types', 'campaign type', 'types', 'type']);
      const productsIndex = getHeaderIndex(headers, ['products', 'product']);
      const notesIndex = getHeaderIndex(headers, ['notes', 'note']);
      const lastChangedActionIndex = getHeaderIndex(headers, ['last change', 'last changed action', 'last action']);
      const lastChangedByIndex = getHeaderIndex(headers, ['last changed by', 'changed by']);
      const lastChangedAtIndex = getHeaderIndex(headers, ['last changed at', 'changed at']);

      const imported = rows
        .slice(1)
        .filter(row => row.some(cell => String(cell).trim() !== ''))
        .map((row, index) => normalizeTask({
          __backendId: generateId(),
          month: getCsvValue(row, monthIndex, 0),
          name: getCsvValue(row, nameIndex, 1) || 'Campaign',
          text: getCsvValue(row, nameIndex, 1) || 'Campaign',
          date: getCsvValue(row, dateIndex, 2),
          status: getCsvValue(row, statusIndex, -1),
          priority: getCsvValue(row, priorityIndex, -1),
          owner: getCsvValue(row, ownerIndex, -1),
          budget: getCsvValue(row, budgetIndex, -1),
          campaignTypes: getCsvValue(row, campaignTypesIndex, -1),
          products: getCsvValue(row, productsIndex, 3),
          notes: getCsvValue(row, notesIndex, 4),
          lastChangedAction: getCsvValue(row, lastChangedActionIndex, -1) || 'Imported',
          lastChangedBy: getCsvValue(row, lastChangedByIndex, -1) || getCurrentActor(),
          lastChangedAt: getCsvValue(row, lastChangedAtIndex, -1) || new Date().toISOString(),
          updatedAt: getCsvValue(row, lastChangedAtIndex, -1) || new Date().toISOString(),
          color: COLORS[index % COLORS.length],
          order: Date.now() + index
        }, index));

      if (!imported.length) throw new Error('No campaign rows were found in the CSV.');

      const shouldReplace = window.confirm('Replace current campaigns with this CSV? Click Cancel to append the imported campaigns.');
      const nextTasks = shouldReplace ? imported : [...allTasks, ...imported];
      refreshTasks(nextTasks);
      syncAllTasksToFirestore();
      showToast(`${imported.length} campaign${imported.length === 1 ? '' : 's'} imported`);
    } catch (error) {
      console.error('CSV import failed:', error);
      showToast(error.message || 'CSV import failed');
    } finally {
      event.target.value = '';
    }
  };

  reader.onerror = () => {
    showToast('Could not read the CSV file');
    event.target.value = '';
  };

  reader.readAsText(file);
}

function initializeFirestoreSync() {
  if (!window.firebase || !firebase.apps) {
    showToast('Firebase SDK not loaded. Using local data.');
    return false;
  }

  try {
    if (!firebase.apps.length) {
      firebase.initializeApp(FIREBASE_CONFIG);
    }

    authInstance = firebase.auth();
    firestoreDb = firebase.firestore();
    updateAuthUi();

    authInstance.onAuthStateChanged(user => {
      currentUser = user;
      updateAuthUi();

      if (firestoreUnsubscribe) {
        firestoreUnsubscribe();
        firestoreUnsubscribe = null;
      }

      if (!currentUser) {
        isFirestoreReady = false;
        allTasks = [];
        setLastSavedNote('');
        renderTasks();
        return;
      }

      isFirestoreReady = true;
      const collection = getFirestoreCollection();
      firestoreUnsubscribe = collection.orderBy('order').onSnapshot(snapshot => {
        isApplyingRemoteSnapshot = true;
        const remoteTasks = snapshot.docs.map((doc, index) => normalizeTask({
          __backendId: doc.id,
          ...doc.data()
        }, index));

        if (remoteTasks.length) {
          allTasks = remoteTasks;
          renderTasks();
          const latestUpdatedAt = remoteTasks
            .map(task => task.updatedAt)
            .filter(Boolean)
            .sort()
            .pop();
          setLastSavedNote(latestUpdatedAt || '');
        } else if (allTasks.length) {
          isApplyingRemoteSnapshot = false;
          syncAllTasksToFirestore();
        } else {
          renderTasks();
        }
        isApplyingRemoteSnapshot = false;
      }, error => {
        isApplyingRemoteSnapshot = false;
        console.error('Firestore live sync failed:', error);
        isFirestoreReady = false;
        showToast('Firestore unavailable');
      });
    });

    return true;
  } catch (error) {
    console.error('Failed to initialize Firebase:', error);
    isFirestoreReady = false;
    showToast('Firebase setup failed. Using local data.');
    return false;
  }
}

async function signInWithGoogle() {
  if (!authInstance) {
    showToast('Firebase Auth is not ready yet');
    return;
  }

  try {
    const provider = new firebase.auth.GoogleAuthProvider();
    await authInstance.signInWithPopup(provider);
    showToast('Signed in with Google');
  } catch (error) {
    console.error('Google sign-in failed:', error);
    showToast(error.code === 'auth/unauthorized-domain'
      ? 'Authorize your GitHub Pages domain in Firebase Auth first'
      : 'Google sign-in failed');
  }
}

async function signOutUser() {
  if (!authInstance) return;
  try {
    await authInstance.signOut();
    showToast('Signed out');
  } catch (error) {
    console.error('Sign-out failed:', error);
    showToast('Could not sign out');
  }
}

// Apply config to UI
function applyConfig(config) {
  const c = { ...defaultConfig, ...config };
  document.body.style.background = c.background_color;
  document.getElementById('calTitle').textContent = c.calendar_title;
  document.getElementById('calTitle').style.color = c.primary_action_color;
  document.getElementById('calTitle').style.fontFamily = `${c.font_family}, 'Metropolis', Arial, sans-serif`;
  document.getElementById('taskCount').style.color = c.secondary_action_color;
  document.getElementById('taskCount').style.fontSize = `${c.font_size}px`;
  document.getElementById('calTitle').style.fontSize = `${c.font_size * 2.2}px`;

  document.querySelectorAll('.month-col').forEach(col => {
    col.style.background = c.surface_color;
    col.querySelector('div').style.color = c.secondary_action_color;
  });
  document.querySelectorAll('.task-card').forEach(card => {
    card.style.color = c.text_color;
  });
  document.querySelectorAll('.add-btn').forEach(btn => {
    btn.style.color = c.secondary_action_color;
    btn.style.borderColor = c.secondary_action_color + '33';
  });
  document.documentElement.style.setProperty('--accent', c.primary_action_color);
}

(async () => {
  buildGrid();
  initEditMonthOptions();
  initEditDateControls();
  initCampaignTypeCheckboxes();
  initFilterOptions();
  initFilterToggle();
  initUpcomingToggle();
  initQuarterToggle();
  initDeletedToggle();
  setLastSavedNote('');
  applyConfig(defaultConfig);
  renderTasks();
  initializeFirestoreSync();
  lucide.createIcons();
})();
