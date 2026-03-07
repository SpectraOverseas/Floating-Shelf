async function hashPassword(password) {
  const encoder = new TextEncoder();
  const data = encoder.encode(password);
  const hashBuffer = await crypto.subtle.digest('SHA-256', data);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map((byte) => byte.toString(16).padStart(2, '0')).join('');
}

async function login() {
  const emailInput = document.getElementById('email');
  const passwordInput = document.getElementById('password');
  const errorMessage = document.getElementById('errorMessage');

  if (!emailInput || !passwordInput) {
    return;
  }

  const email = emailInput.value.trim().toLowerCase();
  const password = passwordInput.value;

  if (errorMessage) {
    errorMessage.textContent = '';
  }

  try {
    const hashedPassword = await hashPassword(password);
    const response = await fetch('users.json', { cache: 'no-store' });

    if (!response.ok) {
      throw new Error('Unable to load users.');
    }

    const users = await response.json();
    const isValidUser = users.some(
      (user) => user.email.toLowerCase() === email && user.password === hashedPassword
    );

    if (!isValidUser) {
      if (errorMessage) {
        errorMessage.textContent = 'Invalid email or password';
      }
      return;
    }

    sessionStorage.setItem('auth', 'true');
    window.location.href = 'index.html';
  } catch (error) {
    if (errorMessage) {
      errorMessage.textContent = 'Invalid email or password';
    }
  }
}

function checkAuth() {
  if (sessionStorage.getItem('auth') !== 'true') {
    window.location.href = 'login.html';
  }
}

function checkLoginAuth() {
  if (sessionStorage.getItem('auth') === 'true') {
    window.location.href = 'index.html';
  }
}

function logout() {
  sessionStorage.removeItem('auth');
  window.location.href = 'login.html';
}
