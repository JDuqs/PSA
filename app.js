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
        const cancelBtn = modalEl.querySelector('.btn-secondary'); // Get cancel button
        
        titleEl.innerText = title;
        textEl.innerText = message;
        
        // Create modal instance
        const modal = new bootstrap.Modal(modalEl, { backdrop: 'static', keyboard: false });
        modal.show();

        // Handle Confirm
        const handleConfirm = () => {
            modal.hide();
            cleanup();
            resolve(true);
        };

        // Handle Cancel
        const handleCancel = () => {
            modal.hide(); // Bootstrap handles removal, but good to be explicit if needed
            cleanup();
            resolve(false);
        };

        // Cleanup listeners to avoid duplicates if modal is reused (though we create new instance here)
        const cleanup = () => {
            yesBtn.onclick = null;
            cancelBtn.onclick = null; // Clear cancel listener too
            modalEl.removeEventListener('hidden.bs.modal', handleCancelModalEvent);
        };

        // Event listener for when modal is hidden via other means (keyboard, backdrop if enabled)
        const handleCancelModalEvent = () => {
            resolve(false);
        };

        yesBtn.onclick = handleConfirm;
        cancelBtn.onclick = handleCancel; // Attach to Cancel button
        
        // Safety net: if they click X or escape (if allowed)
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
            
            if (userError || !userData) throw new Error("Account not found in office records. Please Sign Up first.");
            if (userData.approved !== true) throw new Error("Access Denied: Pending Admin/Guard approval.");
        }

        const { data, error } = await supabase.auth.signInWithPassword({ email, password });
        
        if (error) {
            if (error.message.includes("Email not confirmed")) {
                throw new Error("System Config Error: Please ask Admin to disable 'Confirm Email' in Supabase.");
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
    if (pass !== passConfirm) return alert("Passwords do not match! Please check and try again.");
    
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

        alert("Account created successfully! Please wait for the Admin/Guard to confirm your access.");
        await supabase.auth.signOut(); 
        window.location.href = 'index.html';

    } catch (error) {
        if (error.message.includes("rate limit") || error.status === 429) {
            alert("Signup Limit: Please wait or ask Admin to disable 'Confirm Email'.");
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
        console.warn("loadRegistrationRequests called, but requestTableBody not found. (Not on admin.html?)");
        return;
    }

    if(section) section.style.display = 'block';

    const fetchRequests = async () => {
        console.log("Fetching registration requests...");
        try {
            const { data, error } = await supabase.from('registration_requests').select('*');
            if (error) throw error;
            
            tbody.innerHTML = "";
            if (!data || data.length === 0) {
                tbody.innerHTML = "<tr><td colspan='4' class='text-center text-muted py-3'>No pending requests found.</td></tr>";
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
            tbody.innerHTML = `<tr><td colspan='4' class='text-center text-danger py-3'>Error loading requests. Check console.</td></tr>`;
        }
    };

    fetchRequests();
    supabase.channel('reg-requests').on('postgres_changes', { event: '*', schema: 'public', table: 'registration_requests' }, fetchRequests).subscribe();
}

window.approveUser = async (reqId, email, name, password) => {
    if (!await showConfirm("Approve", `Approve access for ${email}?`)) return;
    try {
        const { error: upsertError } = await supabase.from('users').upsert({
            email: email, name: name, password: password, approved: true, role: 'user'
        }, { onConflict: 'email' });

        if (upsertError) throw upsertError;

        await supabase.from('registration_requests').delete().eq('id', reqId);
        alert("User Approved! Refreshing...");
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
    console.log("Initializing Dashboard for:", user.email);
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
    }
}

// ==========================================
// D. ISSUANCE & CART
// ==========================================
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

// --- INVENTORY AUTOCOMPLETE ---
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

// ==========================================
// E. RECORDS & TABLES & PAGINATION
// ==========================================
function loadAllRecords(user) {
    const activeTbody = document.getElementById('activeTableBody');
    const historyTbody = document.getElementById('historyTableBody');
    const isAdmin = user.email === ADMIN_EMAIL;
    
    const fetchRecords = async () => {
        console.log("Fetching records...");
        
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

            // Initially render both
            renderTable('active');
            renderTable('history');
        } catch (e) {
            console.error("Error fetching records:", e);
        }
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
    if (state.filter) {
        filtered = rawData.filter(item => 
            Object.values(item).some(val => 
                String(val).toLowerCase().includes(state.filter)
            )
        );
    }
    
    const totalItems = filtered.length;
    const totalPages = Math.ceil(totalItems / state.limit);
    if (state.page > totalPages) state.page = Math.max(1, totalPages);
    
    const start = (state.page - 1) * state.limit;
    const paginatedItems = filtered.slice(start, start + state.limit);
    
    tbody.innerHTML = "";
    if (paginatedItems.length === 0) {
        tbody.innerHTML = `<tr><td colspan="13" class="text-center text-muted py-3">No records found.</td></tr>`;
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
                
                // Allow Admin to Edit Due Date Inline
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

// ADMIN FUNCTION: Update Due Date
window.updateDueDate = async (id, newDate) => {
    if (!id || !newDate) return;
    if (!await showConfirm("Update", `Change due date to ${newDate}?`)) {
        window.refreshTableData(); // Revert UI
        return;
    }

    try {
        const { error } = await supabase.from('gate_passes').update({ due_date: newDate }).eq('id', id);
        if (error) throw error;
        alert("Due date updated.");
        window.refreshTableData();
    } catch (e) {
        alert("Update failed: " + e.message);
    }
};

document.getElementById('tableSearch')?.addEventListener('input', (e) => {
    const activeTabButton = document.querySelector('.nav-link.active');
    const targetId = activeTabButton.getAttribute('data-bs-target');
    const type = targetId === '#activeTab' ? 'active' : 'history';
    
    // Update filter for current type
    paginationState[type].filter = e.target.value.toLowerCase();
    paginationState[type].page = 1; 
    renderTable(type);
});

// Return Logic with Manual Refresh
document.getElementById('returnBtn')?.addEventListener('click', async () => {
    const s = document.getElementById('returnSerial').value;
    const g = document.getElementById('guardIn').value;
    if (!s || !g) return alert("Fill fields");

    const { data } = await supabase.from('gate_passes').select('id').eq('serial', s).eq('status', 'OUT').single();
    if (!data) return alert("Not found");

    if (!await showConfirm("Return", `Return ${s}?`)) return;

    await supabase.from('gate_passes').update({ status: 'RETURNED', guard_in: g, time_return: new Date().toISOString() }).eq('id', data.id);
    document.getElementById('returnSerial').value="";
    if(window.refreshTableData) window.refreshTableData();
});
window.selectRow = (s) => document.getElementById('returnSerial').value = s;

// ==========================================
// F. UTILS & EXPORTS
// ==========================================
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

// FIX: Get Selected Items based on EXPORT CONTEXT
function getSelectedItems() {
    const containerId = currentExportContext === 'active' ? 'activeTableBody' : 'historyTableBody';
    const container = document.getElementById(containerId);
    if (!container) return [];
    
    const checkboxes = container.querySelectorAll('.export-check:checked');
    if (!checkboxes.length) { alert("Select items"); return []; }
    return Array.from(checkboxes).map(c => JSON.parse(decodeURIComponent(c.value)));
}

// Export Modals - Sets the context for getSelectedItems
const openExportModal = (context) => {
    currentExportContext = context; // Stores 'active' or 'history'
    const containerId = context === 'active' ? 'activeTableBody' : 'historyTableBody';
    const container = document.getElementById(containerId);
    const count = container ? container.querySelectorAll('.export-check:checked').length : 0;
    
    if(count === 0) return alert(`Select items from ${context === 'active' ? 'Active Records' : 'History Log'}.`);
    
    document.getElementById('exportModalTitle').innerText = `Export ${count} items`;
    new bootstrap.Modal(document.getElementById('exportModal')).show();
}
document.getElementById('openExportModalBtn')?.addEventListener('click', () => openExportModal('active'));
document.getElementById('openExportHistoryModalBtn')?.addEventListener('click', () => openExportModal('history'));

// UPDATED: EXCEL EXPORT (Clean Format based on Table Columns)
document.getElementById('btnExportExcel')?.addEventListener('click', async () => {
    const items = getSelectedItems();
    if (!items.length) return;

    // Check for different borrowers
    const borrowers = new Set(items.map(i => (i.borrower || "").trim().toLowerCase()));
    if (borrowers.size > 1) {
        return alert("Error: You cannot export multiple borrowers to the same file. Please select items for a single borrower only.");
    }

    const borrowerName = items[0].borrower || "Unknown";
    
    // Close modal BEFORE confirming
    const exportModalEl = document.getElementById('exportModal');
    const modalInstance = bootstrap.Modal.getInstance(exportModalEl);
    if (modalInstance) modalInstance.hide();

    if (!await showConfirm("Export Excel", `Generate Excel file for ${borrowerName}?`)) return;

    // Map to nice columns based on which table we are viewing
    const exportData = items.map(item => {
        // Base columns shared by both
        const base = {
            "Gate Pass ID": item.unique_id,
            "Borrower": item.borrower,
            "Description": item.description,
            "Serial No.": item.serial,
            "Property No.": item.property_no || '',
            "Asset Tag": item.asset_no || '',
            "Destination": item.destination,
            "Project": item.project || '',
            "Time Out": item.time_out ? new Date(item.time_out).toLocaleString() : ''
        };

        if (currentExportContext === 'active') {
            return {
                ...base,
                "Guard Out": item.guard_out,
                "Due Date": item.due_date || '',
                "Status": item.status
            };
        } else {
            // History Context
            return {
                ...base,
                "Time Returned": item.time_return ? new Date(item.time_return).toLocaleString() : '',
                "Guard In": item.guard_in,
                "Status": item.status
            };
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

    // Check for different borrowers
    const borrowers = new Set(items.map(i => (i.borrower || "").trim().toLowerCase()));
    if (borrowers.size > 1) {
        return alert("Error: You cannot export multiple borrowers to the same Gate Pass. Please select items for a single borrower only.");
    }

    const borrowerName = items[0].borrower || "Unknown";

    // Close modal BEFORE confirming
    const exportModalEl = document.getElementById('exportModal');
    const modalInstance = bootstrap.Modal.getInstance(exportModalEl);
    if (modalInstance) modalInstance.hide();

    if (!await showConfirm("Export Gate Pass", `Generate Gate Pass PDF for ${borrowerName}?`)) return;

    const groupByBorrower = (arr) => {
        return arr.reduce((acc, obj) => {
            const key = obj.borrower;
            if (!acc[key]) acc[key] = [];
            acc[key].push(obj);
            return acc;
        }, {});
    };

    const groupedItems = groupByBorrower(items);
    const { jsPDF } = window.jspdf;

    Object.keys(groupedItems).forEach((borrowerName) => {
        const doc = new jsPDF();
        const borrowerItems = groupedItems[borrowerName];
        const firstItem = borrowerItems[0];

        // ===== FINAL FOOTER FUNCTION =====
        const stampFooter = () => {
            const pageHeight = doc.internal.pageSize.height;
            const pageWidth = doc.internal.pageSize.width;

            const footerY = pageHeight - 35;

            doc.setFontSize(8);
            doc.setFont("helvetica", "normal");

            doc.setLineWidth(0.3);
            doc.setDrawColor(0, 0, 0);

            doc.line(10, footerY - 5, 200, footerY - 5);

            doc.text("3rd Floor STWLPC Building, 335-338 Sen. Gil Puyat Avenue (Buendia)", 105, footerY, { align: "center" });
            doc.text("Barangay 49 Zone 7, Pasay City Philippines 1300", 105, footerY + 4, { align: "center" });
            doc.text("Telephone (632) 833-8284 • Telefax (632) 834-0051", 105, footerY + 8, { align: "center" });
            doc.text("Email Address: ncr5@psa.gov.ph, Website: www.psa.gov.ph", 105, footerY + 12, { align: "center" });

            const pageNumber = doc.internal.getCurrentPageInfo().pageNumber;
            doc.text(`Page ${pageNumber}`, pageWidth - 20, pageHeight - 10);
        };

        // ===== HEADER =====
        doc.setTextColor(0, 0, 0);

        doc.setFontSize(11);
        doc.text("REPUBLIC OF THE PHILIPPINES", 105, 15, { align: "center" });

        doc.setFontSize(12);
        doc.setFont("helvetica", "bold");
        doc.text("PHILIPPINE STATISTICS AUTHORITY", 105, 20, { align: "center" });

        doc.setFontSize(10);
        doc.setFont("helvetica", "normal");
        doc.text("NCR PROVINCIAL STATISTICAL OFFICE V", 105, 25, { align: "center" });
        doc.text("LAS PIÑAS MUNTINLUPA PARAÑAQUE PASAY", 105, 30, { align: "center" });

        const dateStr = new Date().toLocaleDateString('en-US', {
            month: 'long',
            day: '2-digit',
            year: 'numeric'
        });

        doc.text("Date:", 15, 45);
        doc.text(dateStr, 25, 45);
        doc.line(25, 46, 25 + doc.getTextWidth(dateStr), 46);

        doc.text("GP No.: ", 140, 45);
        const gpVal = firstItem.unique_id;
        doc.text(gpVal, 155, 45);
        doc.rect(153, 40, doc.getTextWidth(gpVal) + 4, 7);

        doc.text("Annex A", 170, 52);

        // ===== BOLD + UNDERLINE STYLED TEXT SECTION =====
        let currentX = 15;
        let currentY = 60;
        const lineHeight = 5;

        function printStyled(text, isBold = false) {
            doc.setFont("helvetica", isBold ? "bold" : "normal");

            const width = doc.getTextWidth(text);

            if (currentX + width > 195) {
                currentX = 15;
                currentY += lineHeight;
            }

            doc.text(text, currentX, currentY);

            if (isBold) {
                doc.setLineWidth(0.5);
                doc.setDrawColor(0, 0, 0);
                doc.line(currentX, currentY + 1, currentX + width, currentY + 1);
            }

            currentX += width;
        }

        printStyled("TO THE GUARD ON DUTY: Please allow ");
        printStyled(firstItem.borrower || "__________", true);
        printStyled(" to bring out property from PSA Office to ");
        printStyled(firstItem.destination || "__________", true);
        printStyled(" for the purpose of ");
        printStyled(firstItem.project || "__________", true);
        printStyled(".");

        currentY += 15;

        const formatTime = (t) => t ? new Date(t).toLocaleString() : '';

        const tableBody = borrowerItems.map(item => [
            item.description,
            item.serial,
            item.property_no,
            item.asset_no,
            item.destination,
            formatTime(item.time_out),
            item.time_return ? formatTime(item.time_return) : ''
        ]);

        // ===== ULTRA DARK BLACK TABLE =====
        doc.autoTable({
            startY: currentY,

            head: [[
                'Description',
                'Serial Number',
                'Property Number',
                'Asset Tag',
                'Destination',
                'Time Out',
                'Time Returned'
            ]],

            body: tableBody,

            theme: 'grid',

            // MODIFIED: White Background, Black Text for Header
            headStyles: {
                fillColor: [255, 255, 255], // White background
                textColor: [0, 0, 0],       // Black text
                lineColor: [0, 0, 0],
                lineWidth: 0.5
            },

            // FORCE BLACK TEXT EVERYWHERE
            styles: {
                textColor: [0, 0, 0],
                lineColor: [0, 0, 0],
                lineWidth: 0.4,
                fontSize: 9
            },

            // BODY CELLS ALSO BLACK
            bodyStyles: {
                textColor: [0, 0, 0]
            },

            margin: {
                bottom: 100
            }
        });

        // ===== SIGNATORIES =====
        let finalY = doc.lastAutoTable.finalY + 15;

        if (finalY > doc.internal.pageSize.height - 120) {
            doc.addPage();
            finalY = 30;
        }

        doc.setFontSize(10);
        doc.setTextColor(0, 0, 0);

        doc.text("Remarks:", 15, finalY);

        finalY += 8;
        doc.line(15, finalY, 195, finalY);
        finalY += 8;
        doc.line(15, finalY, 195, finalY);
        finalY += 8;
        doc.line(15, finalY, 195, finalY);

        finalY += 20;

        doc.text("Checked/Inspected by:", 15, finalY);

        doc.line(15, finalY + 15, 80, finalY + 15);
        doc.setFont("helvetica", "bold");
        doc.text("JENOR B. BLAS", 15, finalY + 20);
        doc.setFont("helvetica", "normal");
        doc.text("Property and Supply Officer", 15, finalY + 25);

        doc.line(100, finalY + 15, 165, finalY + 15);
        doc.setFont("helvetica", "bold");
        doc.text("MARY ANNE G. BASILIO", 100, finalY + 20);
        doc.setFont("helvetica", "normal");
        doc.text("Inspection Officer", 100, finalY + 25);

        finalY += 40;

        if (finalY > doc.internal.pageSize.height - 120) {
            doc.addPage();
            finalY = 30;
        }

        doc.text("Approved by:", 15, finalY);

        doc.line(15, finalY + 15, 80, finalY + 15);
        doc.setFont("helvetica", "bold");
        doc.text("MARICEL M. CARAGAN", 15, finalY + 20);
        doc.setFont("helvetica", "normal");
        doc.text("Supervising Statistical Specialist,", 15, finalY + 25);
        doc.text("Officer-in-Charge, PSA NCR PSO V", 15, finalY + 30);

        // ===== ADD FOOTER TO ALL PAGES LAST =====
        const totalPages = doc.internal.getNumberOfPages();

        for (let i = 1; i <= totalPages; i++) {
            doc.setPage(i);
            stampFooter();
        }

        doc.save(`GatePass_${borrowerName}_${firstItem.unique_id}.pdf`);
    });
});

// ==========================================
// NEW: ACKNOWLEDGEMENT EXPORT
// ==========================================
document.getElementById('btnExportAckReceipt')?.addEventListener('click', async () => {
    const items = getSelectedItems();
    if (items.length === 0) return;

    // Check for different borrowers
    const borrowers = new Set(items.map(i => (i.borrower || "").trim().toLowerCase()));
    if (borrowers.size > 1) {
        return alert("Error: You cannot export multiple borrowers to the same Receipt. Please select items for a single borrower only.");
    }

    const borrowerName = items[0].borrower || "Unknown";

    // Close modal BEFORE confirming
    const exportModalEl = document.getElementById('exportModal');
    const modalInstance = bootstrap.Modal.getInstance(exportModalEl);
    if (modalInstance) modalInstance.hide();

    if (!await showConfirm("Export Receipt", `Generate Acknowledgement Receipt for ${borrowerName}?`)) return;

    // Auto-detect Project Name from the first selected item
    // The user requested this be based on the logs/data
    const projectName = items[0].project || "N/A";

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p', 'mm', 'a4'); // Portrait Mode

    // --- Footer Logic (Same Image Content + Dynamic Page Numbering) ---
    const addFooter = (docInstance) => {
        const pageCount = docInstance.internal.getNumberOfPages();
        const pageWidth = docInstance.internal.pageSize.width;
        const pageHeight = docInstance.internal.pageSize.height;
        const footerY = pageHeight - 20;

        // Line
        docInstance.setLineWidth(0.5);
        docInstance.line(10, footerY - 5, pageWidth - 10, footerY - 5);

        // Image Content Replicated as Text (High Quality)
        docInstance.setFontSize(8);
        docInstance.setFont("helvetica", "normal");
        docInstance.setTextColor(0, 0, 0);

        docInstance.text("3rd Floor STWLPC Building, 335-338 Sen. Gil Puyat Avenue (Buendia)", pageWidth / 2, footerY, { align: "center" });
        docInstance.text("Barangay 49 Zone 7, Pasay City Philippines 1300", pageWidth / 2, footerY + 4, { align: "center" });
        docInstance.text("Telephone (632) 833-8284 • Telefax (632) 834-0051", pageWidth / 2, footerY + 8, { align: "center" });
        docInstance.text("Email Address: ncr5@psa.gov.ph, Website: www.psa.gov.ph", pageWidth / 2, footerY + 12, { align: "center" });
    };

    // --- Header ---
    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    doc.text("Ref No.: 2026-0002", 15, 15); // Example Ref or could be auto-generated

    doc.setFontSize(11);
    doc.text("REPUBLIC OF THE PHILIPPINES", 105, 15, { align: "center" });
    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.text("PHILIPPINE STATISTICS AUTHORITY", 105, 20, { align: "center" });

    doc.setFont("helvetica", "bold");
    doc.setFontSize(14);
    doc.text("Acknowledgment Form", 105, 30, { align: "center" });

    // --- Body Text ---
    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    
    const text1 = "All hired field-based personnel for the specified project listed below acknowledges the receipt of the following: a) tablet, b) accessories compatible case and adapter, and c) powerbank.";
    
    const text2 = "All personnel who were given these devices will be held liable for any acts of negligence and malicious intent resulting to the loss or damage of these tablets. Should there be a lost/damaged tablet, the responsible personnel should immediately inform the incident to their immediate supervisor. Upon the evaluation of the Philippine Statistics Authority (PSA) Provincial Statistical Office (PSO) Chief Statistical Specialist (CSS), an anticipated cost required to repair the damage in the tablet must be shouldered by the liable personnel. In the event that the tablet is lost, a salary deduction equivalent to the market value of the comparable device must be charged against the responsible personnel. Due to this, it is crucial to exercise caution and care to the equipment/device entrusted by the PSA to every field-based personnel for the successful and secure operationalization.";

    const text3 = "Affixing your name and signature in the next page signifies that you hereby acknowledge the receipt of the above-listed devices/items under your name and fully understand the responsibilities attached to these.";

    const splitText1 = doc.splitTextToSize(text1, 180);
    doc.text(splitText1, 15, 40);

    const splitText2 = doc.splitTextToSize(text2, 180);
    doc.text(splitText2, 15, 55);

    const splitText3 = doc.splitTextToSize(text3, 180);
    // Calculate Y based on length of text2 (approx line height 4-5mm)
    let currentY = 55 + (splitText2.length * 5) + 5;
    doc.text(splitText3, 15, currentY);

    currentY += (splitText3.length * 5) + 10;

    // Project Info (Auto-detected)
    doc.setFont("helvetica", "bold");
    doc.text(`Project: ${projectName}`, 15, currentY);
    doc.text("Instructor: ___________________________", 15, currentY + 7);

    currentY += 15;

    // --- Table Data Preparation (Auto-Detection) ---
    const tableData = items.map((item, index) => {
        let accessories = "-";
        const descLower = (item.description || "").toLowerCase();
        
        // Auto-detect "With Powerbank..." based on description
        if (descLower.includes('tablet') || descLower.includes('samsung') || descLower.includes('ipad') || descLower.includes('tab')) {
            accessories = "With Powerbank and/or Accessories";
        } else if (descLower.includes('laptop')) {
            accessories = "With Charger and Bag";
        }

        return [
            index + 1,
            "", // Name of Hired Based Personnel (Left Blank as requested)
            (item.description && item.description.toLowerCase().includes('samsung')) ? "Samsung" : (item.description || ""),
            item.serial,
            item.asset_no || "",
            // Tag Column REMOVED
            accessories,
            "", // Signature (Empty)
            ""  // Date (Empty)
        ];
    });

    // --- Table Generation ---
    doc.autoTable({
        startY: currentY,
        head: [[
            "No.", 
            "Name of Hired\nBased Personnel", 
            "Tablet Brand", 
            "Serial Number", 
            "Asset Tag\nNumber", 
            "With Powerbank\nand/or Accessories", 
            "Signature", 
            "Date of\nAcknowledgement"
        ]],
        body: tableData,
        theme: 'grid',
        headStyles: {
            fillColor: [255, 255, 255],
            textColor: [0, 0, 0],
            lineColor: [0, 0, 0],
            lineWidth: 0.3,
            halign: 'center',
            valign: 'middle',
            fontStyle: 'bold',
            fontSize: 8
        },
        styles: {
            textColor: [0, 0, 0],
            lineColor: [0, 0, 0],
            lineWidth: 0.3,
            fontSize: 8,
            valign: 'middle'
        },
        columnStyles: {
            0: { width: 10, halign: 'center' }, // No
            1: { width: 35 }, // Name
            2: { width: 20 }, // Brand
            3: { width: 25 }, // Serial
            4: { width: 20, halign: 'center' }, // Asset Tag Number (Slightly wider)
            // Tag column removed (index shifted)
            5: { width: 30, fontSize: 7 }, // Accessories (Wider to compensate)
            6: { width: 25 }, // Signature
            7: { width: 20 }  // Date
        },
        didDrawPage: function (data) {
            // Footer on every page
            addFooter(doc);
        }
    });

    // Add Page Numbers (Total Pages)
    const totalPages = doc.internal.getNumberOfPages();
    for (let i = 1; i <= totalPages; i++) {
        doc.setPage(i);
        doc.setFontSize(8);
        doc.text(`Page ${i} of ${totalPages}`, doc.internal.pageSize.width - 25, doc.internal.pageSize.height - 4);
    }

    doc.save(`AcknowledgementReceipt_${projectName}_${new Date().toISOString().split('T')[0]}.pdf`);
});

// NEW: TRANSMITTAL EXPORT FUNCTIONALITY
document.getElementById('btnExportTransmittal')?.addEventListener('click', async () => {
    const items = getSelectedItems();
    if (items.length === 0) return;

    // Check for different borrowers
    const borrowers = new Set(items.map(i => (i.borrower || "").trim().toLowerCase()));
    if (borrowers.size > 1) {
        return alert("Error: You cannot export multiple borrowers to the same Transmittal. Please select items for a single borrower only.");
    }

    const borrowerName = items[0].borrower || "Unknown";

    // Close modal BEFORE confirming
    const exportModalEl = document.getElementById('exportModal');
    const modalInstance = bootstrap.Modal.getInstance(exportModalEl);
    if (modalInstance) modalInstance.hide();

    if (!await showConfirm("Export Transmittal", `Generate Transmittal Form for ${borrowerName}?`)) return;

    // Group items by Destination
    const grouped = items.reduce((acc, item) => {
        const key = item.destination || 'Unspecified';
        if (!acc[key]) acc[key] = [];
        acc[key].push(item);
        return acc;
    }, {});

    const { jsPDF } = window.jspdf;

    Object.keys(grouped).forEach(dest => {
        const doc = new jsPDF();
        const batch = grouped[dest];
        
        // Calculate Summary & Accessories Logic
        const summary = {};
        batch.forEach(item => {
            const d = (item.description || "").toLowerCase();
            let key = "Item";
            
            // Auto-detection logic for Accessories and Summary Count
            if (d.includes('tablet') || d.includes('samsung') || d.includes('galaxy')) {
                // Main Item
                key = "Samsung Tablet";
                summary[key] = (summary[key] || 0) + 1;
                
                // Implicit Accessories (Counted separately in summary)
                summary["Adapter"] = (summary["Adapter"] || 0) + 1;
                summary["Type C Cable"] = (summary["Type C Cable"] || 0) + 1;
                summary["Box"] = (summary["Box"] || 0) + 1;
                
                // Row display
                item._acc = "With type c cable, box and adapter";
            } else if (d.includes('laptop')) {
                key = "Laptop";
                summary[key] = (summary[key] || 0) + 1;
                summary["Charger"] = (summary["Charger"] || 0) + 1;
                summary["Bag"] = (summary["Bag"] || 0) + 1;
                item._acc = "With charger and bag";
            } else {
                key = item.description || "Equipment";
                summary[key] = (summary[key] || 0) + 1;
                item._acc = "-";
            }
        });

        // --- PDF LAYOUT ---
        // Header
        doc.setFontSize(10); doc.setTextColor(0,0,0); doc.text("Republic of the Philippines", 105, 15, {align:'center'});
        doc.setFontSize(11); doc.setFont("helvetica", "bold"); doc.text("PHILIPPINE STATISTICS AUTHORITY", 105, 20, {align:'center'});
        doc.setFontSize(10); doc.setFont("helvetica", "normal"); 
        doc.text("2024 POPCEN-CBMS", 105, 25, {align:'center'});
        doc.text(dest.toUpperCase(), 105, 30, {align:'center'});
        
        doc.setFontSize(12); doc.setFont("helvetica", "bold"); doc.text("TRANSMITTAL / RECEIPT FORM", 105, 40, {align:'center'});
        doc.setFontSize(9); doc.setFont("helvetica", "italic"); doc.text("(Accomplish in duplicate copies)", 105, 45, {align:'center'});

        // Summary Section (Top Left) - FULL WIDTH
        const summaryData = Object.entries(summary).map(([k, v]) => [k, v]);
        doc.autoTable({
            startY: 55,
            head: [['ITEMS', 'QTY']],
            body: summaryData,
            theme: 'grid',
            styles: { 
                fontSize: 9, 
                textColor: [0,0,0], 
                lineColor: [0,0,0], 
                lineWidth: 0.4 
            },
            headStyles: { 
                fillColor: [220, 220, 220], 
                textColor: [0,0,0], 
                fontStyle: 'bold', 
                halign: 'center', 
                lineColor: [0,0,0], 
                lineWidth: 0.4 
            },
            columnStyles: { 0: { halign: 'left' }, 1: { width: 30, halign: 'center' } }, // First column stretches, second fixed width
        });

        // Main Table starts after summary
        let mainY = doc.lastAutoTable.finalY + 10;

        const tableBody = batch.map((item, i) => [
            i + 1,
            `${item.description}\n\n${item.serial}`, // Combined Stacked Item/Serial
            item.asset_no || '',
            '1',
            item._acc
        ]);

        doc.autoTable({
            startY: mainY,
            head: [['No.', 'SERIAL No. / ITEM NAME', 'ASSET TAG No.', 'UNIT', 'ACCESSORIES']],
            body: tableBody,
            theme: 'grid',
            headStyles: { 
                fillColor: [255, 255, 255], 
                textColor: [0,0,0], 
                lineColor: [0,0,0], 
                lineWidth: 0.4, 
                halign: 'center',
                valign: 'middle',
                fontStyle: 'bold'
            },
            styles: { 
                textColor: [0,0,0], 
                lineColor: [0,0,0], 
                lineWidth: 0.4,
                valign: 'middle',
                cellPadding: 3
            },
            columnStyles: {
                0: { halign: 'center', width: 10 },
                1: { width: 80 },
                2: { halign: 'center' },
                3: { halign: 'center', width: 15 },
                4: { width: 50 }
            }
        });

        // Signatories Section - SIDE BY SIDE BUT WITH DETAILED BLOCK STRUCTURE
        let finalY = doc.lastAutoTable.finalY + 20;

        // Check page break once for both
        if (finalY + 50 > doc.internal.pageSize.height - 20) {
            doc.addPage();
            finalY = 40;
        }

        doc.setFontSize(10);
        doc.setTextColor(0,0,0);
        
        // Helper to draw a full signature block at a specific X
        const drawSigBlock = (x, label) => {
            let currentBlockY = finalY; // Start from same Y for both
            
            doc.text(label + ":", x, currentBlockY);
            currentBlockY += 15; // Space before line
            
            doc.setLineWidth(0.5);
            doc.setDrawColor(0,0,0);
            doc.line(x, currentBlockY, x + 80, currentBlockY); // Signature Line
            doc.setFontSize(8);
            doc.text("SIGNATURE OVER PRINTED NAME", x + 40, currentBlockY + 5, {align:'center'});
            
            currentBlockY += 15; // Space for next line
            doc.line(x, currentBlockY, x + 80, currentBlockY); // Position Line
            doc.text("POSITION / DESIGNATION", x + 40, currentBlockY + 5, {align:'center'});

            currentBlockY += 15; // Space for next line
            doc.line(x, currentBlockY, x + 80, currentBlockY); // Date Line
            doc.text("DATE SIGNED", x + 40, currentBlockY + 5, {align:'center'});
            
            doc.setFontSize(10);
        };

        // Draw Left Block
        drawSigBlock(15, "Transmitted by");
        // Draw Right Block
        drawSigBlock(115, "Received by");

        // ===== ADD PAGE NUMBERS TO ALL PAGES =====
        const totalPages = doc.internal.getNumberOfPages();
        for (let i = 1; i <= totalPages; i++) {
            doc.setPage(i);
            doc.setFontSize(8);
            doc.setTextColor(0, 0, 0);
            doc.text(`Page ${i} of ${totalPages}`, doc.internal.pageSize.width - 25, doc.internal.pageSize.height - 10);
        }

        doc.save(`Transmittal_${dest}_${new Date().toISOString().split('T')[0]}.pdf`);
    });
});

// Import Logic
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
            
            bulkImportData = json.map(row => ({
                serial: row['serial_no'] || row['Serial'],
                description: row['description'] || row['Description'],
                asset_no: row['asset_no'] || row['Asset'],
                property_no: row['property_no'] || row['Property']
            })).filter(x => x.serial);

            const tbody = document.getElementById('importBody');
            tbody.innerHTML = "";
            bulkImportData.slice(0, 5).forEach(d => {
                tbody.innerHTML += `<tr><td>${d.serial}</td><td>${d.property_no}</td><td>${d.description}</td><td>${d.asset_no}</td></tr>`;
            });
            document.getElementById('importPreview').style.display='block';
            document.getElementById('saveBulkBtn').style.display='block';
        };
        r.readAsArrayBuffer(f);
    });
}
document.getElementById('saveBulkBtn')?.addEventListener('click', async () => {
    if(!bulkImportData.length) return;
    if(!await showConfirm("Import", `Save ${bulkImportData.length} items?`)) return;
    
    const { error } = await supabase.from('inventory').upsert(bulkImportData, { onConflict: 'serial' });
    if(error) alert("Error: " + error.message);
    else { alert("Imported!"); bulkImportData = []; document.getElementById('importPreview').style.display='none'; }
});

// Wait for DOM before initializing
window.addEventListener('DOMContentLoaded', () => {
    checkUserSession();
});