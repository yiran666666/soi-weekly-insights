/**
 * Shared user identity module.
 * Prompts for name on first visit, stores in localStorage.
 * Provides helpers for ownership checks and API headers.
 */

const AUTH_KEY = 'gds_hub_username';

// Get stored username (or null)
function getCurrentUser() {
  return localStorage.getItem(AUTH_KEY);
}

// Set username
function setCurrentUser(name) {
  localStorage.setItem(AUTH_KEY, name.trim());
}

// Check if the current user owns an item (case-insensitive match)
function isOwner(itemAuthor) {
  const user = getCurrentUser();
  if (!user || !itemAuthor) return false;
  return user.toLowerCase() === itemAuthor.toLowerCase();
}

// Headers to include with edit/delete requests
function authHeaders() {
  return { 'X-Author': getCurrentUser() || '' };
}

// Prompt user for their name if not set. Returns a promise that resolves with the username.
function ensureUser() {
  return new Promise((resolve) => {
    const existing = getCurrentUser();
    if (existing) { resolve(existing); return; }
    showNamePrompt(resolve);
  });
}

// Build and show the name prompt modal
function showNamePrompt(onDone) {
  const overlay = document.createElement('div');
  overlay.id = 'auth-overlay';
  overlay.innerHTML = `
    <div id="auth-modal">
      <h2>Welcome to China GDS Hub</h2>
      <p>Enter your name to get started. This identifies your contributions so only you can edit or delete them.</p>
      <input type="text" id="auth-name-input" placeholder="e.g. Yiran" autofocus>
      <button id="auth-save-btn">Continue</button>
    </div>
  `;

  // Inject styles
  const style = document.createElement('style');
  style.textContent = `
    #auth-overlay {
      position: fixed; inset: 0; background: rgba(0,0,0,0.45); z-index: 9999;
      display: flex; align-items: center; justify-content: center;
      backdrop-filter: blur(6px); -webkit-backdrop-filter: blur(6px);
    }
    #auth-modal {
      background: #fff; border-radius: 18px; padding: 36px; width: 400px; max-width: 90vw;
      box-shadow: 0 20px 60px rgba(0,0,0,0.25); text-align: center;
      font-family: -apple-system, BlinkMacSystemFont, 'SF Pro Display', 'Helvetica Neue', sans-serif;
    }
    #auth-modal h2 { font-size: 22px; font-weight: 700; color: #1d1d1f; margin-bottom: 8px; }
    #auth-modal p { font-size: 14px; color: #86868b; line-height: 1.5; margin-bottom: 20px; }
    #auth-name-input {
      width: 100%; padding: 12px 16px; border: 1px solid #d2d2d7; border-radius: 10px;
      font-size: 16px; font-family: inherit; outline: none; margin-bottom: 16px; text-align: center;
    }
    #auth-name-input:focus { border-color: #0071e3; box-shadow: 0 0 0 3px rgba(0,113,227,0.1); }
    #auth-save-btn {
      width: 100%; padding: 12px; border-radius: 10px; border: none;
      background: #0071e3; color: #fff; font-size: 15px; font-weight: 600;
      cursor: pointer; font-family: inherit; transition: background 0.15s;
    }
    #auth-save-btn:hover { background: #0062cc; }
    #auth-save-btn:disabled { background: #d2d2d7; cursor: not-allowed; }
  `;
  document.head.appendChild(style);
  document.body.appendChild(overlay);

  const input = document.getElementById('auth-name-input');
  const btn = document.getElementById('auth-save-btn');

  function submit() {
    const name = input.value.trim();
    if (!name) { input.focus(); return; }
    setCurrentUser(name);
    overlay.remove();
    style.remove();
    onDone(name);
  }

  btn.addEventListener('click', submit);
  input.addEventListener('keydown', (e) => { if (e.key === 'Enter') submit(); });
}

// Render a small "Logged in as" indicator. Call after ensureUser().
function renderUserBadge(containerSelector) {
  const user = getCurrentUser();
  if (!user) return;
  const container = document.querySelector(containerSelector);
  if (!container) return;

  const badge = document.createElement('div');
  badge.className = 'user-identity-badge';
  badge.innerHTML = `
    <span class="uid-name">${escapeHtmlAuth(user)}</span>
    <button class="uid-change" onclick="changeUser()" title="Switch user">Change</button>
  `;
  container.appendChild(badge);

  if (!document.getElementById('uid-styles')) {
    const s = document.createElement('style');
    s.id = 'uid-styles';
    s.textContent = `
      .user-identity-badge {
        display: flex; align-items: center; gap: 8px; padding: 8px 16px;
        font-size: 12px; color: #86868b; border-top: 0.5px solid #d2d2d7; margin-top: 8px;
      }
      .uid-name { font-weight: 600; color: #1d1d1f; }
      .uid-change {
        background: none; border: none; color: #0071e3; font-size: 12px;
        cursor: pointer; padding: 0; font-family: inherit;
      }
      .uid-change:hover { text-decoration: underline; }
    `;
    document.head.appendChild(s);
  }
}

// Allow switching user
window.changeUser = function() {
  localStorage.removeItem(AUTH_KEY);
  location.reload();
};

function escapeHtmlAuth(str) {
  const d = document.createElement('div');
  d.textContent = str || '';
  return d.innerHTML;
}
