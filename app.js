// 1. SUPABASE CONFIGURATION
const SUPABASE_URL = 'https://hkgjyrtjemdditazycim.supabase.co';
const SUPABASE_KEY = 'sb_publishable_MmYypzfD2NV8vUi8GEmbRQ_OpZJcHYN';
const supabase = window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY);

// GLOBAL VARIABLES
const ADMIN_EMAIL = "admin@psa.gov.ph";
let cart = [];
let currentUserName = "";
let bulkImportData = [];
let inventoryCache = {};
let allInventoryItems = []; 
let borrowedSerials = new Set(); 
let currentUser = null; 
let currentExportContext = 'active'; 

// --- PAGINATION STATE ---
let activeData = [];
let historyData = [];
let paginationState = {
    active: { page: 1, limit: 10, filter: '' },
    history: { page: 1, limit: 10, filter: '' }
};

console.log("App.js loaded. Initializing...");

// ==========================================
// UTILS
// ==========================================
function showConfirm(title, message) {
    return new Promise((resolve) => {
        const modalEl = document.getElementById('confirmModal');
        if (!modalEl) return resolve(confirm(message));

        const titleEl = document.getElementById('confirmTitle');
        const textEl = document.getElementById('confirmText');
        const yesBtn = document.getElementById('confirmBtnYes');
        const cancelBtn = modalEl.querySelector('.btn-secondary'); 
        
        titleEl.innerText = title;
        textEl.innerText = message;
        
        const modal = new bootstrap.Modal(modalEl, { backdrop: 'static', keyboard: false });
        modal.show();

        const handleConfirm = () => {
            modal.hide();
            cleanup();
            resolve(true);
        };

        const handleCancel = () => {
            modal.hide();
            cleanup();
            resolve(false);
        };

        const cleanup = () => {
            yesBtn.onclick = null;
            if(cancelBtn) cancelBtn.onclick = null;
            modalEl.removeEventListener('hidden.bs.modal', handleCancelModalEvent);
        };

        const handleCancelModalEvent = () => {
            resolve(false);
        };

        yesBtn.onclick = handleConfirm;
        if(cancelBtn) cancelBtn.onclick = handleCancel;
        modalEl.addEventListener('hidden.bs.modal', handleCancelModalEvent, { once: true });
    });
}

function generateGatePassID() {
    const now = new Date();
    const yymmdd = now.toISOString().slice(2, 10).replace(/-/g, '');
    const random = Math.floor(1000 + Math.random() * 9000);
    return `GP-${yymmdd}-${random}`;
}

// ==========================================
// A. AUTHENTICATION & SESSION HANDLING
// ==========================================
const loginForm = document.getElementById('loginForm');
const loginBtn = document.getElementById('loginBtn');
const logoutBtn = document.getElementById('logoutBtn');

async function checkUserSession() {
    const { data: { session } } = await supabase.auth.getSession();
    const path = window.location.pathname;

    console.log("Checking session. Path:", path);

    if (session) {
        currentUser = session.user; 
        console.log("User logged in:", currentUser.email);
        
        // Redirect from public pages to dashboard
        if (path.includes('index.html') || path.includes('signup.html') || path === '/' || path.endsWith('/')) {
            window.location.href = 'dashboard.html';
            return;
        }
        
        // Initialize Dashboard Logic
        if (path.includes('dashboard')) {
            initDashboard(currentUser);
        }

        // Initialize Admin Logic
        if (path.includes('admin')) {
            if (currentUser.email !== ADMIN_EMAIL) {
                // Not admin? Kick to dashboard
                window.location.href = 'dashboard.html';
            } else {
                // Is admin? Load data
                loadRegistrationRequests();
            }
        }

    } else {
        // No Session - Kick to login if on protected pages
        if (path.includes('dashboard') || path.includes('admin')) {
            window.location.href = 'index.html';
        }
    }
}

async function handleLogin(e) {
    if (e) e.preventDefault();
    const email = document.getElementById('email').value.toLowerCase().trim();
    const password = document.getElementById('password').value;

    if (!email || !password) return alert("Please enter email and password");

    const btn = document.getElementById('loginBtn');
    if(btn) { btn.disabled = true; btn.innerText = "Verifying..."; }

    try {
        if (email !== ADMIN_EMAIL) {
            const { data: userData, error: userError } = await supabase
                .from('users')
                .select('*') 
                .eq('email', email)
                .single();
            
            if (userError || !userData) throw new Error("Account not found. Please Sign Up.");
            if (userData.approved !== true) throw new Error("Access Denied: Pending Admin approval.");
        }

        const { data, error } = await supabase.auth.signInWithPassword({ email, password });
        
        if (error) {
            if (error.message.includes("Email not confirmed")) {
                throw new Error("System Config Error: Ask Admin to disable 'Confirm Email'.");
            }
            throw error;
        }

        window.location.href = 'dashboard.html';

    } catch (error) {
        const msgEl = document.getElementById('errorMsg');
        let displayMsg = "Login Failed: " + error.message;
        if (error.message.includes("Invalid login credentials")) displayMsg = "Incorrect email or password.";
        
        if (msgEl) msgEl.innerText = displayMsg;
        else alert(displayMsg);
        
        if(btn) { btn.disabled = false; btn.innerText = "LOGIN TO SYSTEM"; }
    }
}

if (loginBtn) loginBtn.addEventListener('click', handleLogin);
if (loginForm) loginForm.addEventListener('submit', handleLogin);

if (logoutBtn) {
    logoutBtn.addEventListener('click', async () => {
        await supabase.auth.signOut();
        window.location.href = 'index.html';
    });
}

// ==========================================
// B. REGISTRATION & APPROVAL
// ==========================================
const requestBtn = document.getElementById('requestBtn');
const signupForm = document.getElementById('signupForm');

async function handleSignup(e) {
    if (e) e.preventDefault();
    
    const firstName = document.getElementById('regFirstName').value.trim();
    const lastName = document.getElementById('regLastName').value.trim();
    const email = document.getElementById('regEmail').value.toLowerCase().trim();
    const pass = document.getElementById('regPass').value;
    const passConfirm = document.getElementById('regPassConfirm').value;

    if (!firstName || !lastName || !email) return alert("Please fill all details.");
    if (pass.length < 6) return alert("Password must be at least 6 characters.");
    if (pass !== passConfirm) return alert("Passwords do not match!");
    
    const name = `${firstName} ${lastName}`;
    const btn = document.getElementById('requestBtn');
    if(btn) { btn.disabled = true; btn.innerText = "Processing..."; }

    try {
        const { data: authData, error: authError } = await supabase.auth.signUp({
            email, password: pass, options: { data: { full_name: name } }
        });
        if (authError) throw authError;

        await supabase.from('registration_requests').insert([{ name, email, pass, status: 'PENDING' }]);
        
        const uid = authData.user ? authData.user.uid : null;
        const { error: userError } = await supabase.from('users').insert([{ 
            email, name, password: pass, approved: false, role: 'user', uid: uid
        }]);

        if (userError) console.error("Profile warning:", userError);

        alert("Request submitted! Wait for Admin approval.");
        await supabase.auth.signOut(); 
        window.location.href = 'index.html';

    } catch (error) {
        if (error.message.includes("rate limit") || error.status === 429) {
            alert("Signup Limit: Please wait or ask Admin.");
        } else {
            alert("Error: " + error.message);
        }
        if(btn) { btn.disabled = false; btn.innerText = "SUBMIT ACCESS REQUEST"; }
    }
}

if (requestBtn) requestBtn.addEventListener('click', handleSignup);
if (signupForm) signupForm.addEventListener('submit', handleSignup);

async function loadRegistrationRequests() {
    const tbody = document.getElementById('requestTableBody');
    const section = document.getElementById('adminRequestSection');
    
    if (!tbody) {
        console.warn("requestTableBody not found. Not on admin page?");
        return;
    }

    if(section) section.style.display = 'block';

    const fetchRequests = async () => {
        try {
            const { data, error } = await supabase.from('registration_requests').select('*');
            if (error) throw error;
            
            tbody.innerHTML = "";
            if (!data || data.length === 0) {
                tbody.innerHTML = "<tr><td colspan='4' class='text-center text-muted py-3'>No pending requests.</td></tr>";
                return;
            }
            data.forEach(d => {
                tbody.innerHTML += `
                    <tr>
                        <td>${d.name}</td>
                        <td>${d.email}</td>
                        <td class="text-muted"><i>Hidden</i></td>
                        <td class="text-center">
                            <button class="btn btn-success btn-sm me-1" onclick="window.approveUser('${d.id}', '${d.email}', '${d.name}', '${d.pass}')">Confirm</button>
                            <button class="btn btn-danger btn-sm" onclick="window.cancelRequest('${d.id}')">Deny</button>
                        </td>
                    </tr>`;
            });
        } catch (err) {
            console.error("Error loading requests:", err);
            tbody.innerHTML = `<tr><td colspan='4' class='text-center text-danger'>Error loading data.</td></tr>`;
        }
    };

    fetchRequests();
    supabase.channel('reg-requests').on('postgres_changes', { event: '*', schema: 'public', table: 'registration_requests' }, fetchRequests).subscribe();
}

window.approveUser = async (reqId, email, name, password) => {
    if (!await showConfirm("Approve", `Approve ${email}?`)) return;
    try {
        const { error: upsertError } = await supabase.from('users').upsert({
            email: email, name: name, password: password, approved: true, role: 'user'
        }, { onConflict: 'email' });

        if (upsertError) throw upsertError;

        await supabase.from('registration_requests').delete().eq('id', reqId);
        alert("User Approved!");
        window.location.reload(); 
    } catch (e) { alert(e.message); }
};

window.cancelRequest = async (id) => {
    if (await showConfirm("Reject", "Delete this request?")) {
        try {
            await supabase.from('registration_requests').delete().eq('id', id);
            alert("Request rejected.");
            window.location.reload();
        } catch (e) { alert(e.message); }
    }
};

// ==========================================
// C. DASHBOARD INITIALIZATION
// ==========================================
async function initDashboard(user) {
    const borrowerInput = document.getElementById('borrower');
    if (!user || !user.email) return;

    try {
        const { data: userProfile } = await supabase.from('users').select('name').eq('email', user.email).single();
        if (userProfile) currentUserName = userProfile.name;
        else currentUserName = user.email.split('@')[0];
    } catch (e) {
        currentUserName = user.email.split('@')[0];
    }

    const userDisplay = document.getElementById('currentUserDisplay');
    if (userDisplay) {
        userDisplay.innerHTML = `<i class="fa fa-circle-user me-2" style="color: var(--psa-yellow);"></i><span class="fw-bold">${currentUserName}</span>`;
    }

    if (user.email === ADMIN_EMAIL) {
        const adminBtn = document.getElementById('adminPanelBtn');
        const bulkNav = document.getElementById('bulkImportNav');
        if (adminBtn) adminBtn.style.display = 'inline-block';
        if (bulkNav) bulkNav.style.display = 'block';
        
        if (borrowerInput) {
            borrowerInput.readOnly = false;
            borrowerInput.style.backgroundColor = '#fff';
            borrowerInput.placeholder = "Enter Borrower's Name";
        }
    } else {
        if (window.location.href.includes('admin')) {
            window.location.href = 'dashboard.html';
            return;
        }
        if (borrowerInput) {
            borrowerInput.value = currentUserName;
            borrowerInput.readOnly = true;
            borrowerInput.style.backgroundColor = '#e9ecef';
        }
    }
    
    if(document.getElementById('activeTableBody')) {
        loadAllRecords(user);
        loadInventory(); 
        updateClock();
        initSearchListeners(); // Initialize new search logic
    }
}

function updateClock() {
    const timeEl = document.getElementById('clockTime');
    const dateEl = document.getElementById('clockDate');
    if(!timeEl) return;
    setInterval(() => {
        const now = new Date();
        dateEl.innerText = now.toDateString();
        timeEl.innerText = now.toLocaleTimeString();
    }, 1000);
}

document.getElementById('addToCartBtn')?.addEventListener('click', () => {
    const s = document.getElementById('serial').value;
    const p = document.getElementById('propertyNum').value;
    const d = document.getElementById('desc').value;
    const a = document.getElementById('asset').value;

    if (!s || !d) return alert("Fill Serial/Desc");
    if (cart.some(i => i.serial === s)) return alert("Already in list");

    if (borrowedSerials.has(s)) {
        return alert(`Item ${s} is currently MARKED AS OUT. Return it first.`);
    }

    cart.push({ serial: s, property_no: p, desc: d, asset: a });
    renderCart();
    refreshInventoryDropdown(); 
    
    document.getElementById('serial').value=""; 
    document.getElementById('propertyNum').value=""; 
    document.getElementById('desc').value=""; 
    document.getElementById('asset').value="";
});

function renderCart() {
    const tbody = document.getElementById('cartTableBody');
    if (!tbody) return;
    tbody.innerHTML = cart.length === 0 ? '<tr><td colspan="5" class="text-center text-muted">Empty</td></tr>' : "";
    cart.forEach((item, i) => {
        tbody.innerHTML += `<tr><td>${item.serial}</td><td>${item.property_no||'-'}</td><td>${item.desc}</td><td>${item.asset}</td><td class="text-center"><button onclick="window.removeItem(${i})" class="btn btn-sm btn-danger py-0">&times;</button></td></tr>`;
    });
    document.getElementById('issueBtn').disabled = cart.length === 0;
}
window.removeItem = (i) => { 
    cart.splice(i, 1); 
    renderCart(); 
    refreshInventoryDropdown(); 
}

document.getElementById('issueBtn')?.addEventListener('click', async () => {
    const borrower = document.getElementById('borrower').value;
    const guard = document.getElementById('guardOut').value;
    const dest = document.getElementById('destination').value;
    const proj = document.getElementById('project').value;
    const due = document.getElementById('dueDate').value;

    if (!borrower || !guard || !dest || !due) return alert("Fill all required fields.");

    for (const item of cart) {
        const { data } = await supabase.from('gate_passes').select('*').eq('serial', item.serial).eq('status', 'OUT');
        if (data && data.length > 0) return alert(`Serial ${item.serial} is already OUT.`);
    }

    if (!await showConfirm("Confirm", `Issue ${cart.length} items?`)) return;

    try {
        const { data: { user } } = await supabase.auth.getUser();
        const batchID = generateGatePassID();
        
        const records = cart.map(item => ({
            unique_id: batchID,
            issuer_email: user.email,
            borrower, guard_out: guard, destination: dest, project: proj, due_date: due,
            serial: item.serial, property_no: item.property_no, description: item.desc, asset_no: item.asset,
            time_out: new Date(), status: "OUT", time_return: null
        }));

        const { error } = await supabase.from('gate_passes').insert(records);
        if (error) throw error;

        alert("Issued!");
        cart = []; renderCart();
        refreshInventoryDropdown(); 
        if (user.email === ADMIN_EMAIL) document.getElementById('borrower').value = "";
        window.refreshTableData();
    } catch(e) { alert(e.message); }
});

async function loadInventory() {
    let list = document.getElementById('inventoryList');
    if (!list) {
        list = document.createElement('datalist');
        list.id = 'inventoryList';
        document.body.appendChild(list);
        document.getElementById('serial')?.setAttribute('list', 'inventoryList');
    }
    
    const { data } = await supabase.from('inventory').select('*').limit(2000);
    if(data) {
        allInventoryItems = data; 
        inventoryCache = {};
        data.forEach(i => inventoryCache[i.serial] = i);
        updateBorrowedStatus();
    }
}

async function updateBorrowedStatus() {
    const { data } = await supabase.from('gate_passes').select('serial').eq('status', 'OUT');
    if(data) {
        borrowedSerials = new Set(data.map(d => d.serial));
        refreshInventoryDropdown();
    }
}

function refreshInventoryDropdown() {
    const list = document.getElementById('inventoryList');
    if (!list) return;
    list.innerHTML = "";
    
    const cartSet = new Set(cart.map(c => c.serial));
    
    allInventoryItems.forEach(i => {
        if (!borrowedSerials.has(i.serial) && !cartSet.has(i.serial)) {
            const opt = document.createElement('option');
            opt.value = i.serial;
            opt.innerText = i.description; 
            list.appendChild(opt);
        }
    });
}

document.getElementById('serial')?.addEventListener('change', (e) => {
    const item = inventoryCache[e.target.value];
    if(item) {
        document.getElementById('desc').value = item.description;
        document.getElementById('asset').value = item.asset_no;
        if(document.getElementById('propertyNum')) document.getElementById('propertyNum').value = item.property_no;
    }
});

function loadAllRecords(user) {
    const isAdmin = user.email === ADMIN_EMAIL;
    
    const fetchRecords = async () => {
        try {
            let query = supabase.from('gate_passes').select('*').eq('status', 'OUT').order('time_out', { ascending: false }).limit(1000);
            if (!isAdmin) query = query.eq('issuer_email', user.email);
            const { data: aData, error: aError } = await query;
            if (aError) throw aError;
            if(aData) activeData = aData;

            let hQuery = supabase.from('gate_passes').select('*').eq('status', 'RETURNED').order('time_return', { ascending: false }).limit(1000);
            if (!isAdmin) hQuery = hQuery.eq('issuer_email', user.email);
            const { data: hData, error: hError } = await hQuery;
            if (hError) throw hError;
            if(hData) historyData = hData;

            // Initially render both with empty search
            renderTable('active');
            renderTable('history');
        } catch (e) { console.error("Error fetching records:", e); }
    };

    window.refreshTableData = () => {
        fetchRecords();
        updateBorrowedStatus();
    };

    fetchRecords();
    supabase.channel('gp_updates').on('postgres_changes', { event: '*', schema: 'public', table: 'gate_passes' }, () => {
        fetchRecords(); updateBorrowedStatus(); 
    }).subscribe();
}

function renderTable(type) {
    const state = paginationState[type];
    const rawData = type === 'active' ? activeData : historyData;
    const tbody = document.getElementById(type === 'active' ? 'activeTableBody' : 'historyTableBody');
    
    if (!tbody) return;
    
    let filtered = rawData;
    
    // === REVISED SEARCH LOGIC ===
    if (state.filter) {
        const term = state.filter.toLowerCase();
        filtered = rawData.filter(item => {
            // Prepare formatted dates for search visibility
            const formattedTimeOut = item.time_out ? new Date(item.time_out).toLocaleString() : '';
            const formattedTimeReturn = item.time_return ? new Date(item.time_return).toLocaleString() : '';
            const formattedDueDate = item.due_date ? item.due_date : ''; // usually already string YYYY-MM-DD

            // Search visible fields only
            const fields = [
                item.unique_id, item.borrower, item.description, 
                item.serial, item.property_no, item.asset_no, 
                item.destination, item.project, item.guard_out, item.guard_in,
                formattedTimeOut, formattedTimeReturn, formattedDueDate, item.status
            ];
            return fields.some(val => val && String(val).toLowerCase().includes(term));
        });
    }
    
    const totalItems = filtered.length;
    const totalPages = Math.ceil(totalItems / state.limit);
    if (state.page > totalPages) state.page = Math.max(1, totalPages);
    
    const start = (state.page - 1) * state.limit;
    const paginatedItems = filtered.slice(start, start + state.limit);
    
    tbody.innerHTML = "";
    if (paginatedItems.length === 0) {
        tbody.innerHTML = `<tr><td colspan="13" class="text-center text-muted py-3">No records found matching "${state.filter}".</td></tr>`;
    } else {
        const today = new Date().toISOString().split('T')[0];
        const isAdmin = currentUser && currentUser.email === ADMIN_EMAIL;
        
        paginatedItems.forEach(data => {
            const safeData = encodeURIComponent(JSON.stringify(data));
            
            if (type === 'active') {
                let badge = '<span class="badge bg-primary">OUT</span>';
                if (data.due_date) {
                    if (data.due_date < today) { badge = '<span class="badge bg-danger">OVERDUE</span>'; }
                    else if (data.due_date === today) { badge = '<span class="badge bg-warning text-dark">DUE TODAY</span>'; }
                }
                
                let dueDateCell = data.due_date || '-';
                if (isAdmin) {
                    dueDateCell = `<input type="date" class="form-control form-control-sm border-warning" 
                                    style="min-width:130px; font-size: 0.8rem;"
                                    value="${data.due_date || ''}" 
                                    onchange="window.updateDueDate('${data.id}', this.value)" 
                                    onclick="event.stopPropagation()">`;
                }

                tbody.innerHTML += `
                <tr onclick="window.selectRow('${data.serial}')" style="cursor:pointer">
                    <td onclick="event.stopPropagation()"><input type="checkbox" class="export-check" value="${safeData}"></td>
                    <td class="fw-bold text-primary">${data.unique_id}</td>
                    <td>${data.borrower}</td><td>${data.description}</td><td>${data.serial}</td><td>${data.property_no||'-'}</td><td>${data.asset_no}</td>
                    <td>${data.destination}</td><td>${data.project}</td>
                    <td>${new Date(data.time_out).toLocaleString()}</td><td>${data.guard_out}</td>
                    <td onclick="event.stopPropagation()">${dueDateCell}</td>
                    <td>${badge}</td>
                </tr>`;
            } else {
                tbody.innerHTML += `
                <tr>
                    <td onclick="event.stopPropagation()"><input type="checkbox" class="export-check" value="${safeData}"></td>
                    <td>${data.unique_id}</td><td>${data.borrower}</td><td>${data.description}</td><td>${data.serial}</td><td>${data.property_no||'-'}</td>
                    <td>${new Date(data.time_out).toLocaleString()}</td>
                    <td class="text-success fw-bold">${new Date(data.time_return).toLocaleString()}</td>
                    <td>${data.guard_in}</td><td>${data.destination}</td><td>${data.project}</td>
                </tr>`;
            }
        });
    }

    renderPaginationControls(type, totalItems, totalPages);
    if(type === 'active') {
        let overdueCount = activeData.filter(d => d.due_date && d.due_date < new Date().toISOString().split('T')[0]).length;
        const alertBox = document.getElementById('overdueAlert');
        if (alertBox) {
            if (overdueCount > 0) {
                alertBox.innerText = `${overdueCount} ITEM(S) OVERDUE`;
                alertBox.className = "alert alert-danger text-center fw-bold";
            } else {
                alertBox.innerText = "ALL ON SCHEDULE";
                alertBox.className = "alert alert-success text-center fw-bold";
            }
        }
    }
    updateSelectionCount();
}

function renderPaginationControls(type, totalItems, totalPages) {
    const containerId = type + 'Pagination';
    let container = document.getElementById(containerId);
    if (!container) {
        const tableDiv = document.getElementById(type === 'active' ? 'activeTableBody' : 'historyTableBody').closest('.table-responsive');
        container = document.createElement('div');
        container.id = containerId;
        container.className = "d-flex justify-content-between align-items-center mt-3 pt-2 border-top";
        tableDiv.after(container);
    }
    const state = paginationState[type];
    container.innerHTML = `
        <div class="d-flex align-items-center gap-2"><span class="small text-muted">Show</span>
            <select class="form-select form-select-sm" style="width:70px" onchange="changeLimit('${type}', this.value)">
                <option value="5" ${state.limit==5?'selected':''}>5</option><option value="10" ${state.limit==10?'selected':''}>10</option>
                <option value="50" ${state.limit==50?'selected':''}>50</option><option value="1000" ${state.limit==1000?'selected':''}>All</option>
            </select>
            <span class="small text-muted">Total: ${totalItems}</span>
        </div>
        <div class="btn-group">
            <button class="btn btn-sm btn-outline-secondary" onclick="changePage('${type}', -1)" ${state.page===1?'disabled':''}>Prev</button>
            <button class="btn btn-sm btn-outline-secondary disabled">Page ${state.page} / ${totalPages||1}</button>
            <button class="btn btn-sm btn-outline-secondary" onclick="changePage('${type}', 1)" ${state.page>=totalPages?'disabled':''}>Next</button>
        </div>`;
}

window.changeLimit = (type, limit) => { paginationState[type].limit = parseInt(limit); paginationState[type].page = 1; renderTable(type); };
window.changePage = (type, dir) => { paginationState[type].page += dir; renderTable(type); };

window.updateDueDate = async (id, newDate) => {
    if (!id || !newDate) return;
    if (!await showConfirm("Update", `Change due date to ${newDate}?`)) {
        window.refreshTableData();
        return;
    }
    try {
        const { error } = await supabase.from('gate_passes').update({ due_date: newDate }).eq('id', id);
        if (error) throw error;
        alert("Due date updated.");
        window.refreshTableData();
    } catch (e) { alert("Update failed: " + e.message); }
};

// === SEARCH LISTENER INITIALIZATION ===
function initSearchListeners() {
    // 1. Active Log Search (Global Search Bar)
    const activeSearch = document.getElementById('tableSearch');
    if (activeSearch) {
        // Update functionality to ONLY filter Active Data
        activeSearch.placeholder = "Search Active Records...";
        
        // Remove potential duplicate listeners by recreating the element (safest without named functions)
        const newActiveSearch = activeSearch.cloneNode(true);
        activeSearch.parentNode.replaceChild(newActiveSearch, activeSearch);
        
        newActiveSearch.addEventListener('input', (e) => {
            const term = e.target.value.toLowerCase();
            paginationState.active.filter = term;
            paginationState.active.page = 1; 
            renderTable('active');
        });

        // 2. Tab Switching Visuals for Active Search
        const tabEls = document.querySelectorAll('button[data-bs-toggle="pill"], button[data-bs-toggle="tab"]');
        tabEls.forEach(tabEl => {
            tabEl.addEventListener('shown.bs.tab', (event) => {
                const targetId = event.target.getAttribute('data-bs-target');
                const isHistory = targetId && targetId.includes('history');
                
                // Dim/Disable active search when not on active tab to reduce confusion
                newActiveSearch.disabled = isHistory;
                newActiveSearch.style.opacity = isHistory ? '0.4' : '1';
                newActiveSearch.placeholder = isHistory ? "Use History Search below" : "Search Active Records...";
            });
        });
    }

    // 3. Inject Dedicated History Search Bar
    const historyTbody = document.getElementById('historyTableBody');
    if (historyTbody && !document.getElementById('historySearchInput')) {
        const tableResponsive = historyTbody.closest('.table-responsive');
        if (tableResponsive) {
            // Create container
            const searchContainer = document.createElement('div');
            searchContainer.className = "d-flex justify-content-end mb-2 align-items-center";
            searchContainer.innerHTML = `
                <div class="input-group" style="width: 300px;">
                    <span class="input-group-text bg-white border-end-0"><i class="fa fa-search text-muted"></i></span>
                    <input type="text" id="historySearchInput" class="form-control border-start-0" placeholder="Search History Log...">
                </div>
            `;
            
            // Insert before the table container
            tableResponsive.parentNode.insertBefore(searchContainer, tableResponsive);

            // Bind Listener
            const historySearch = document.getElementById('historySearchInput');
            historySearch.addEventListener('input', (e) => {
                const term = e.target.value.toLowerCase();
                paginationState.history.filter = term;
                paginationState.history.page = 1;
                renderTable('history');
            });
        }
    }
}

document.getElementById('returnBtn')?.addEventListener('click', async () => {
    const s = document.getElementById('returnSerial').value;
    const g = document.getElementById('guardIn').value;
    if (!s || !g) return alert("Fill fields");

    const { data } = await supabase.from('gate_passes').select('id').eq('serial', s).eq('status', 'OUT').single();
    if (!data) return alert("Not found");

    if (!await showConfirm("Return", `Return ${s}?`)) return;

    await supabase.from('gate_passes').update({ status: 'RETURNED', guard_in: g, time_return: new Date().toISOString() }).eq('id', data.id);
    document.getElementById('returnSerial').value="";
    // Updated refresh call
    if(window.refreshTableData) window.refreshTableData();
});
window.selectRow = (s) => document.getElementById('returnSerial').value = s;

function updateSelectionCount() {
    const count = document.querySelectorAll('.export-check:checked').length;
    const l1 = document.getElementById('selectionCount');
    const l2 = document.getElementById('historySelectionCount');
    if (l1) l1.innerText = `${count} selected`;
    if (l2) l2.innerText = `${count} selected`;
}
document.addEventListener('change', (e) => { if(e.target.classList.contains('export-check')) updateSelectionCount(); });
const toggleAll = (chk, tableBodyId) => {
    const table = document.getElementById(tableBodyId);
    if(table) {
        table.querySelectorAll('.export-check').forEach(c => c.checked = chk);
        updateSelectionCount();
    }
};
document.getElementById('selectAllBtn')?.addEventListener('click', () => toggleAll(true, 'activeTableBody'));
document.getElementById('deselectAllBtn')?.addEventListener('click', () => toggleAll(false, 'activeTableBody'));
document.getElementById('selectAllHistoryBtn')?.addEventListener('click', () => toggleAll(true, 'historyTableBody'));
document.getElementById('deselectAllHistoryBtn')?.addEventListener('click', () => toggleAll(false, 'historyTableBody'));

function getSelectedItems() {
    const containerId = currentExportContext === 'active' ? 'activeTableBody' : 'historyTableBody';
    const container = document.getElementById(containerId);
    if (!container) return [];
    
    const checkboxes = container.querySelectorAll('.export-check:checked');
    if (!checkboxes.length) { alert("Select items"); return []; }
    return Array.from(checkboxes).map(c => JSON.parse(decodeURIComponent(c.value)));
}

const openExportModal = (context) => {
    currentExportContext = context;
    const containerId = context === 'active' ? 'activeTableBody' : 'historyTableBody';
    const container = document.getElementById(containerId);
    const count = container ? container.querySelectorAll('.export-check:checked').length : 0;
    
    if(count === 0) return alert(`Select items from ${context === 'active' ? 'Active Records' : 'History Log'}.`);
    
    document.getElementById('exportModalTitle').innerText = `Export ${count} items`;
    new bootstrap.Modal(document.getElementById('exportModal')).show();
}
document.getElementById('openExportModalBtn')?.addEventListener('click', () => openExportModal('active'));
document.getElementById('openExportHistoryModalBtn')?.addEventListener('click', () => openExportModal('history'));

// HELPER: Close Export Modal before confirming
const closeExportAndConfirm = async (actionName, message) => {
    const exportModalEl = document.getElementById('exportModal');
    const modalInstance = bootstrap.Modal.getInstance(exportModalEl) || new bootstrap.Modal(exportModalEl);
    if(modalInstance) modalInstance.hide();
    
    // Small delay to ensure backdrop clears before next modal
    await new Promise(r => setTimeout(r, 150));
    return await showConfirm(actionName, message);
};

document.getElementById('btnExportExcel')?.addEventListener('click', async () => {
    const items = getSelectedItems();
    if (!items.length) return;

    const borrowers = new Set(items.map(i => (i.borrower || "").trim().toLowerCase()));
    if (borrowers.size > 1) return alert("Error: Select items for a single borrower only.");

    const borrowerName = items[0].borrower || "Unknown";
    
    if (!await closeExportAndConfirm("Export Excel", `Generate Excel file for ${borrowerName}?`)) return;

    const exportData = items.map(item => {
        const base = {
            "Gate Pass ID": item.unique_id, "Borrower": item.borrower, "Description": item.description,
            "Serial No.": item.serial, "Property No.": item.property_no || '', "Asset Tag": item.asset_no || '',
            "Destination": item.destination, "Project": item.project || '', "Time Out": item.time_out ? new Date(item.time_out).toLocaleString() : ''
        };
        if (currentExportContext === 'active') {
            return { ...base, "Guard Out": item.guard_out, "Due Date": item.due_date || '', "Status": item.status };
        } else {
            return { ...base, "Time Returned": item.time_return ? new Date(item.time_return).toLocaleString() : '', "Guard In": item.guard_in, "Status": item.status };
        }
    });

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Logs");
    XLSX.writeFile(wb, `PSA_Logs_${currentExportContext}_${new Date().toISOString().split('T')[0]}.xlsx`);
});

document.getElementById('btnExportGatePass')?.addEventListener('click', async () => {
    const items = getSelectedItems();
    if (items.length === 0) return;

    const borrowers = new Set(items.map(i => (i.borrower || "").trim().toLowerCase()));
    if (borrowers.size > 1) return alert("Error: Select items for a single borrower only.");

    const borrowerName = items[0].borrower || "Unknown";
    if (!await closeExportAndConfirm("Export Gate Pass", `Generate Gate Pass PDF for ${borrowerName}?`)) return;

    const groupedItems = { [borrowerName]: items };
    const { jsPDF } = window.jspdf;

    Object.keys(groupedItems).forEach((borrowerName) => {
        const doc = new jsPDF();
        const borrowerItems = groupedItems[borrowerName];
        const firstItem = borrowerItems[0];

        const stampFooter = () => {
            const pageHeight = doc.internal.pageSize.height;
            const pageWidth = doc.internal.pageSize.width;
            const footerY = pageHeight - 35;
            doc.setFontSize(8); doc.setFont("helvetica", "normal"); doc.setLineWidth(0.3); doc.setDrawColor(0, 0, 0);
            doc.line(10, footerY - 5, 200, footerY - 5);
            doc.text("3rd Floor STWLPC Building, 335-338 Sen. Gil Puyat Avenue (Buendia)", 105, footerY, { align: "center" });
            doc.text("Barangay 49 Zone 7, Pasay City Philippines 1300", 105, footerY + 4, { align: "center" });
            doc.text("Telephone (632) 833-8284 • Telefax (632) 834-0051", 105, footerY + 8, { align: "center" });
            doc.text("Email Address: ncr5@psa.gov.ph, Website: www.psa.gov.ph", 105, footerY + 12, { align: "center" });
            doc.text(`Page ${doc.internal.getCurrentPageInfo().pageNumber}`, pageWidth - 20, pageHeight - 10);
        };

        doc.setTextColor(0, 0, 0); doc.setFontSize(11);
        doc.text("REPUBLIC OF THE PHILIPPINES", 105, 15, { align: "center" });
        doc.setFontSize(12); doc.setFont("helvetica", "bold");
        doc.text("PHILIPPINE STATISTICS AUTHORITY", 105, 20, { align: "center" });
        doc.setFontSize(10); doc.setFont("helvetica", "normal");
        doc.text("NCR PROVINCIAL STATISTICAL OFFICE V", 105, 25, { align: "center" });
        doc.text("LAS PIÑAS MUNTINLUPA PARAÑAQUE PASAY", 105, 30, { align: "center" });

        const dateStr = new Date().toLocaleDateString('en-US', { month: 'long', day: '2-digit', year: 'numeric' });
        doc.text("Date:", 15, 45); doc.text(dateStr, 25, 45); doc.line(25, 46, 25 + doc.getTextWidth(dateStr), 46);
        doc.text("GP No.: ", 140, 45);
        doc.text(firstItem.unique_id, 155, 45); doc.rect(153, 40, doc.getTextWidth(firstItem.unique_id) + 4, 7);
        doc.text("Annex A", 170, 52);

        let currentX = 15, currentY = 60;
        function printStyled(text, isBold = false) {
            doc.setFont("helvetica", isBold ? "bold" : "normal");
            const width = doc.getTextWidth(text);
            if (currentX + width > 195) { currentX = 15; currentY += 5; }
            doc.text(text, currentX, currentY);
            if (isBold) { doc.setLineWidth(0.5); doc.line(currentX, currentY + 1, currentX + width, currentY + 1); }
            currentX += width;
        }

        printStyled("TO THE GUARD ON DUTY: Please allow "); printStyled(firstItem.borrower || "__________", true);
        printStyled(" to bring out property from PSA Office to "); printStyled(firstItem.destination || "__________", true);
        printStyled(" for the purpose of "); printStyled(firstItem.project || "__________", true); printStyled(".");
        currentY += 15;

        const tableBody = borrowerItems.map(item => [
            item.description, item.serial, item.property_no, item.asset_no, item.destination,
            item.time_out ? new Date(item.time_out).toLocaleString() : '', item.time_return ? new Date(item.time_return).toLocaleString() : ''
        ]);

        doc.autoTable({
            startY: currentY,
            head: [['Description', 'Serial Number', 'Property Number', 'Asset Tag', 'Destination', 'Time Out', 'Time Returned']],
            body: tableBody, theme: 'grid',
            headStyles: { fillColor: [255, 255, 255], textColor: [0, 0, 0], lineColor: [0, 0, 0], lineWidth: 0.5 },
            styles: { textColor: [0, 0, 0], lineColor: [0, 0, 0], lineWidth: 0.4, fontSize: 9 },
            margin: { bottom: 100 }
        });

        let finalY = doc.lastAutoTable.finalY + 15;
        if (finalY > doc.internal.pageSize.height - 120) { doc.addPage(); finalY = 30; }

        doc.setFontSize(10); doc.setTextColor(0, 0, 0); doc.text("Remarks:", 15, finalY);
        finalY += 8; doc.line(15, finalY, 195, finalY); finalY += 8; doc.line(15, finalY, 195, finalY); finalY += 8; doc.line(15, finalY, 195, finalY);
        finalY += 20; doc.text("Checked/Inspected by:", 15, finalY);
        doc.line(15, finalY + 15, 80, finalY + 15); doc.setFont("helvetica", "bold"); doc.text("JENOR B. BLAS", 15, finalY + 20); doc.setFont("helvetica", "normal"); doc.text("Property and Supply Officer", 15, finalY + 25);
        doc.line(100, finalY + 15, 165, finalY + 15); doc.setFont("helvetica", "bold"); doc.text("MARY ANNE G. BASILIO", 100, finalY + 20); doc.setFont("helvetica", "normal"); doc.text("Inspection Officer", 100, finalY + 25);
        
        finalY += 40; if (finalY > doc.internal.pageSize.height - 120) { doc.addPage(); finalY = 30; }
        
        doc.text("Approved by:", 15, finalY);
        doc.line(15, finalY + 15, 80, finalY + 15); doc.setFont("helvetica", "bold"); doc.text("MARICEL M. CARAGAN", 15, finalY + 20); doc.setFont("helvetica", "normal"); doc.text("Supervising Statistical Specialist,", 15, finalY + 25); doc.text("Officer-in-Charge, PSA NCR PSO V", 15, finalY + 30);

        const totalPages = doc.internal.getNumberOfPages();
        for (let i = 1; i <= totalPages; i++) { doc.setPage(i); stampFooter(); }
        doc.save(`GatePass_${borrowerName}_${firstItem.unique_id}.pdf`);
    });
});

const processImportBtn = document.getElementById('processImportBtn');
const saveBulkBtn = document.getElementById('saveBulkBtn');

if (processImportBtn) {
    processImportBtn.addEventListener('click', () => {
        const f = document.getElementById('inventoryFile').files[0];
        if(!f) return alert("Select file");
        const r = new FileReader();
        r.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const wb = XLSX.read(data, {type: 'array'});
            const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
            bulkImportData = json.map(row => ({ serial: row['serial_no'] || row['Serial'], description: row['description'] || row['Description'], asset_no: row['asset_no'] || row['Asset'], property_no: row['property_no'] || row['Property'] })).filter(x => x.serial);
            const tbody = document.getElementById('importBody');
            tbody.innerHTML = "";
            bulkImportData.slice(0, 5).forEach(d => { tbody.innerHTML += `<tr><td>${d.serial}</td><td>${d.property_no}</td><td>${d.description}</td><td>${d.asset_no}</td></tr>`; });
            document.getElementById('importPreview').style.display='block';
            document.getElementById('saveBulkBtn').style.display='block';
        };
        r.readAsArrayBuffer(f);
    });
}
if(saveBulkBtn) {
    saveBulkBtn.addEventListener('click', async () => {
        if(!bulkImportData.length) return;
        if(!await showConfirm("Import", `Save ${bulkImportData.length} items?`)) return;
        const { error } = await supabase.from('inventory').upsert(bulkImportData, { onConflict: 'serial' });
        if(error) alert("Error: " + error.message);
        else { alert("Imported!"); bulkImportData = []; document.getElementById('importPreview').style.display='none'; }
    });
}

// CRITICAL: Ensure initialization only after DOM is ready
window.addEventListener('DOMContentLoaded', () => {
    checkUserSession();
});