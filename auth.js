/**
 * Shared user identity module.
 * Reads name from login cookie (set at /login page).
 * Falls back to localStorage for backwards compat.
 */

const AUTH_KEY = 'gds_hub_username';

// Get current user — from cookie first, then localStorage
function getCurrentUser() {
  // Try cookie
  const match = document.cookie.match(/gds_user=([^;]+)/);
  if (match) {
    const name = decodeURIComponent(match[1]);
    localStorage.setItem(AUTH_KEY, name); // sync to localStorage
    return name;
  }
  return localStorage.getItem(AUTH_KEY);
}

function isOwner(itemAuthor) {
  const user = getCurrentUser();
  if (!user || !itemAuthor) return false;
  return user.toLowerCase() === itemAuthor.toLowerCase();
}

function authHeaders() {
  return { 'X-Author': getCurrentUser() || '' };
}

// No prompt needed — IAP authenticates the user and the server sets the
// gds_user cookie from the signed-in email. Just resolve with current user.
function ensureUser() {
  return new Promise((resolve) => {
    resolve(getCurrentUser());
  });
}

// Render "Logged in as" badge
function renderUserBadge(containerSelector) {
  const user = getCurrentUser();
  if (!user) return;
  const container = document.querySelector(containerSelector);
  if (!container) return;

  const badge = document.createElement('div');
  badge.className = 'user-identity-badge';
  badge.innerHTML = `
    <span class="uid-name">${escapeHtmlAuth(user)}</span>
    <button class="uid-logout" onclick="logout()" title="Sign out">Sign out</button>
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
      .uid-logout {
        background: none; border: none; color: #0071e3; font-size: 12px;
        cursor: pointer; padding: 0; font-family: inherit;
      }
      .uid-logout:hover { text-decoration: underline; }
    `;
    document.head.appendChild(s);
  }
}

window.logout = function() {
  document.cookie = 'gds_auth=;path=/;max-age=0';
  document.cookie = 'gds_user=;path=/;max-age=0';
  localStorage.removeItem(AUTH_KEY);
  // Sign out of Google IAP itself; it will re-prompt on next visit.
  window.location.href = '/_gcp_iap/clear_login_cookie';
};

function escapeHtmlAuth(str) {
  const d = document.createElement('div');
  d.textContent = str || '';
  return d.innerHTML;
}
