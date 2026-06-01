// kilowatt-auth.js - Client-side authentication for static/dynamic hosting
(function() {
  // Automatically redirect incorrect /public/ URLs to root level
  if (window.location.pathname.includes('/public/')) {
    const cleanPath = window.location.pathname.replace('/public/', '/');
    window.location.href = window.location.origin + cleanPath + window.location.search + window.location.hash;
    return;
  }

  // Check login state via API on page load (excluding preview page)
  if (window.location.pathname.includes('preview.html')) {
    return;
  }

  // Prevent Flash of Unprotected Content (FOUC)
  const isDemo = window.location.search.includes('projectId=demo');
  if (!isDemo) {
    document.documentElement.style.display = 'none';
  }

  async function checkAuth() {
    if (isDemo) {
      document.documentElement.style.display = '';
      return;
    }

    try {
      const res = await fetch('/api/auth/me');
      const data = await res.json();
      if (!data.authenticated) {
        showLoginScreen();
      } else {
        localStorage.setItem('kilowatt_auth_session', data.username);
        setupAuthenticatedUI();
      }
    } catch (err) {
      // Offline / Server down fallback to preview
      showLoginScreen();
    }
  }

  if (document.readyState === 'loading') {
    window.addEventListener('DOMContentLoaded', checkAuth);
  } else {
    checkAuth();
  }

  window.addEventListener('pageshow', (event) => {
    if (event.persisted) {
      window.location.reload();
    }
  });

  function showLoginScreen() {
    // Inject clean styling for login card matching DataLens light palette
    const style = document.createElement('style');
    style.innerHTML = `
      body {
        margin: 0;
        padding: 0;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        background-color: #f4f6fa;
        height: 100vh;
        display: flex;
        align-items: center;
        justify-content: center;
      }
      .auth-container {
        width: 100%;
        max-width: 380px;
        background: #ffffff;
        border: 1px solid #dcdfe6;
        border-radius: 8px;
        box-shadow: 0 4px 16px rgba(31, 36, 51, 0.04);
        padding: 28px 32px;
        box-sizing: border-box;
      }
      .auth-header {
        text-align: center;
        margin-bottom: 20px;
      }
      .auth-logo {
        font-size: 22px;
        font-weight: 800;
        color: #1f2433;
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 8px;
        margin-bottom: 6px;
      }
      .auth-logo::before {
        content: '';
        width: 4px;
        height: 18px;
        background: #4b73ff;
        border-radius: 2px;
      }
      .auth-subtitle {
        font-size: 13px;
        color: #7f8a9e;
      }
      .auth-tabs {
        display: flex;
        border-bottom: 1px solid #e2e5ec;
        margin-bottom: 20px;
      }
      .auth-tab {
        flex: 1;
        text-align: center;
        padding: 8px 0;
        font-size: 13px;
        font-weight: 600;
        color: #7f8a9e;
        cursor: pointer;
        border-bottom: 2px solid transparent;
        transition: all 0.15s ease;
      }
      .auth-tab.active {
        color: #4b73ff;
        border-bottom-color: #4b73ff;
      }
      .auth-form-group {
        display: flex;
        flex-direction: column;
        gap: 6px;
        margin-bottom: 14px;
      }
      .auth-form-group label {
        font-size: 10px;
        font-weight: 700;
        text-transform: uppercase;
        color: #7f8a9e;
        letter-spacing: 0.05em;
      }
      .auth-input {
        padding: 8px 12px;
        font-size: 13px;
        border: 1px solid #dcdfe6;
        border-radius: 4px;
        outline: none;
        transition: border-color 0.15s ease;
        background: #ffffff;
        color: #1f2433;
        font-family: inherit;
      }
      .auth-input:focus {
        border-color: #4b73ff;
      }
      .auth-btn {
        width: 100%;
        padding: 10px;
        font-size: 13px;
        font-weight: 600;
        color: #ffffff;
        background: #4b73ff;
        border: none;
        border-radius: 4px;
        cursor: pointer;
        transition: background 0.15s ease;
        margin-top: 6px;
        font-family: inherit;
      }
      .auth-btn:hover {
        background: #365de0;
      }
      .auth-error {
        color: #d94f5c;
        font-size: 12px;
        margin-top: 10px;
        text-align: center;
        display: none;
        font-weight: 500;
      }
      .auth-hint {
        font-size: 11px;
        color: #7f8a9e;
        text-align: center;
        margin-top: 16px;
        line-height: 1.4;
        background: #f8fafc;
        padding: 8px;
        border-radius: 4px;
        border: 1px solid #eef0f3;
      }
    `;
    document.head.appendChild(style);

    // Overwrite body with authorization container
    document.body.innerHTML = `
      <div class="auth-container">
        <div class="auth-header">
          <div class="auth-logo">Аналитика</div>
          <div class="auth-subtitle">Управление проектами и дашбордами</div>
        </div>
        <div class="auth-tabs">
          <div class="auth-tab active" id="tab-login" onclick="switchAuthTab('login')">Вход</div>
          <div class="auth-tab" id="tab-register" onclick="switchAuthTab('register')">Регистрация</div>
        </div>
        <form id="auth-form" onsubmit="handleAuthSubmit(event)">
          <div class="auth-form-group">
            <label for="auth-username">Логин</label>
            <input type="text" id="auth-username" required class="auth-input" autocomplete="username">
          </div>
          <div class="auth-form-group">
            <label for="auth-password">Пароль</label>
            <input type="password" id="auth-password" required class="auth-input" autocomplete="current-password">
          </div>
          <button type="submit" class="auth-btn" id="auth-submit-btn">Войти</button>
          <div class="auth-error" id="auth-error-msg"></div>
        </form>
        <div class="auth-hint" id="auth-hint-text">
          Стандартный вход:<br>Логин: <b>admin</b> | Пароль: <b>kilowatt2026</b>
        </div>
      </div>
    `;

    let currentTab = 'login';
    window.switchAuthTab = (tab) => {
      currentTab = tab;
      const tabLogin = document.getElementById('tab-login');
      const tabReg = document.getElementById('tab-register');
      const btn = document.getElementById('auth-submit-btn');
      const hint = document.getElementById('auth-hint-text');
      const errMsg = document.getElementById('auth-error-msg');
      errMsg.style.display = 'none';

      if (tab === 'login') {
        tabLogin.classList.add('active');
        tabReg.classList.remove('active');
        btn.innerText = 'Войти';
        hint.style.display = 'block';
      } else {
        tabReg.classList.add('active');
        tabLogin.classList.remove('active');
        btn.innerText = 'Зарегистрироваться';
        hint.style.display = 'none';
      }
    };

    window.handleAuthSubmit = async (e) => {
      e.preventDefault();
      const usernameInput = document.getElementById('auth-username').value.trim();
      const passwordInput = document.getElementById('auth-password').value;
      const errMsg = document.getElementById('auth-error-msg');

      if (!usernameInput || !passwordInput) {
        errMsg.innerText = 'Заполните все поля!';
        errMsg.style.display = 'block';
        return;
      }

      const endpoint = currentTab === 'login' ? '/api/auth/login' : '/api/auth/register';
      
      try {
        const res = await fetch(endpoint, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ username: usernameInput, password: passwordInput })
        });
        const data = await res.json();
        
        if (res.ok && data.status === 'success') {
          localStorage.setItem('kilowatt_auth_session', usernameInput);
          window.location.reload();
        } else {
          errMsg.innerText = data.message || 'Ошибка авторизации';
          errMsg.style.display = 'block';
        }
      } catch (err) {
        errMsg.innerText = 'Не удалось соединиться с сервером';
        errMsg.style.display = 'block';
      }
    };

    // Restore html styling to display
    document.documentElement.style.display = '';
  }

  function setupAuthenticatedUI() {
    document.documentElement.style.display = '';

    const username = localStorage.getItem('kilowatt_auth_session') || 'admin';
    
    // Update avatars and user profile panels
    const userNames = document.querySelectorAll('.user-name');
    userNames.forEach(el => { el.innerText = username; });

    const userAvatars = document.querySelectorAll('.avatar');
    userAvatars.forEach(el => { el.innerText = username.charAt(0).toUpperCase(); });
  }
})();
