// kilowatt-auth.js - Client-side authentication for static/dynamic hosting
(function() {
  // Automatically redirect incorrect /public/ URLs to root level
  if (window.location.pathname.includes('/public/')) {
    const cleanPath = window.location.pathname.replace('/public/', '/');
    window.location.href = window.location.origin + cleanPath + window.location.search + window.location.hash;
    return;
  }

  // Prevent Flash of Unprotected Content (FOUC)
  const hasSession = localStorage.getItem('kilowatt_auth_session');
  if (!hasSession) {
    document.documentElement.style.display = 'none';
  }

  // Prepopulate default user if user database doesn't exist
  if (!localStorage.getItem('kilowatt_users')) {
    const defaultUsers = {
      'admin': 'kilowatt2026'
    };
    localStorage.setItem('kilowatt_users', JSON.stringify(defaultUsers));
  }

  function init() {
    const currentSession = localStorage.getItem('kilowatt_auth_session');
    if (!currentSession) {
      showLoginScreen();
    } else {
      setupAuthenticatedUI();
    }
  }

  if (document.readyState === 'loading') {
    window.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

  // Handle back-forward cache (bfcache)
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
        background-color: #f2f3f7;
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
        box-shadow: 0 4px 16px rgba(31, 36, 51, 0.06);
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
          <div class="auth-logo">КИЛОВАТТ</div>
          <div class="auth-subtitle">Система сквозной аналитики</div>
        </div>
        <div class="auth-tabs">
          <div class="auth-tab active" id="tab-login" onclick="switchAuthTab('login')">Вход</div>
          <div class="auth-tab" id="tab-register" onclick="switchAuthTab('register')">Регистрация</div>
        </div>
        <form id="auth-form" onsubmit="handleAuthSubmit(event)">
          <div class="auth-form-group">
            <label for="auth-username">Имя пользователя (Логин)</label>
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
          Используйте стандартный вход:<br>Логин: <b>admin</b> | Пароль: <b>kilowatt2026</b>
        </div>
      </div>
    `;

    // Expose helpers globally so HTML onclicks/onsubmits function correctly
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

    window.handleAuthSubmit = (e) => {
      e.preventDefault();
      const usernameInput = document.getElementById('auth-username').value.trim();
      const passwordInput = document.getElementById('auth-password').value;
      const errMsg = document.getElementById('auth-error-msg');

      if (!usernameInput || !passwordInput) {
        errMsg.innerText = 'Заполните все поля!';
        errMsg.style.display = 'block';
        return;
      }

      const users = JSON.parse(localStorage.getItem('kilowatt_users')) || {};

      if (currentTab === 'login') {
        if (users[usernameInput] && users[usernameInput] === passwordInput) {
          localStorage.setItem('kilowatt_auth_session', usernameInput);
          window.location.reload();
        } else {
          errMsg.innerText = 'Неверное имя пользователя или пароль';
          errMsg.style.display = 'block';
        }
      } else {
        if (users[usernameInput]) {
          errMsg.innerText = 'Пользователь с таким именем уже существует';
          errMsg.style.display = 'block';
        } else {
          users[usernameInput] = passwordInput;
          localStorage.setItem('kilowatt_users', JSON.stringify(users));
          localStorage.setItem('kilowatt_auth_session', usernameInput);
          window.location.reload();
        }
      }
    };

    // Restore html styling to display
    document.documentElement.style.display = '';
  }

  function setupAuthenticatedUI() {
    // Show page content
    document.documentElement.style.display = '';

    const username = localStorage.getItem('kilowatt_auth_session');
    
    // Update avatar and username display in the header
    const userNames = document.querySelectorAll('.user-name');
    userNames.forEach(el => {
      el.innerText = username;
    });

    const userAvatars = document.querySelectorAll('.avatar');
    userAvatars.forEach(el => {
      el.innerText = username.charAt(0).toUpperCase();
    });

    // Create Logout Button
    const logoutBtn = document.createElement('button');
    logoutBtn.className = 'btn-action';
    logoutBtn.style.padding = '4px 8px';
    logoutBtn.style.fontSize = '12px';
    logoutBtn.style.marginLeft = '12px';
    logoutBtn.style.color = '#d94f5c';
    logoutBtn.style.borderColor = 'rgba(217, 79, 92, 0.2)';
    logoutBtn.style.height = '28px';
    logoutBtn.style.display = 'inline-flex';
    logoutBtn.style.alignItems = 'center';
    logoutBtn.style.gap = '4px';
    logoutBtn.innerHTML = '🚪 Выйти';
    logoutBtn.onclick = () => {
      localStorage.removeItem('kilowatt_auth_session');
      window.location.reload();
    };

    // Try to append Logout Button to different layout places
    const topBarActions = document.querySelector('.topbar-actions');
    const controls = document.querySelector('.controls');
    const userProfile = document.querySelector('.user-profile');

    if (userProfile && topBarActions) {
      // In main index page, place it next to user profile
      topBarActions.appendChild(logoutBtn);
    } else if (userProfile && controls) {
      // In dashboards, place it next to the profile block we injected
      userProfile.parentNode.insertBefore(logoutBtn, userProfile.nextSibling);
    } else if (controls) {
      controls.appendChild(logoutBtn);
    }
  }
})();
