// Get the form elements
const loginForm = document.getElementById('loginForm');
const emailInput = document.getElementById('email');
const passwordInput = document.getElementById('password');
const submitBtn = document.getElementById('submitBtn');
const errorMessage = document.getElementById('errorMessage');

// Handle form submission
loginForm.addEventListener('submit', async function(e) {
    e.preventDefault(); // Don't refresh page
    
    // Get email and password values
    const email = emailInput.value.trim();
    const password = passwordInput.value;
    
    // Reset error message
    errorMessage.style.display = 'none';
    errorMessage.innerHTML = '';
    
    // Validate input
    if (!email || !password) {
        showError('Please enter email and password');
        return;
    }
    
    // Check email format
    if (!isValidEmail(email)) {
        showError('Please enter a valid email address');
        return;
    }
    
    // Show loading state
    submitBtn.disabled = true;
    submitBtn.classList.add('loading');
    submitBtn.innerHTML = 'Logging in...';
    
    try {
        // Send login request to server
        const response = await fetch('/login', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                email: email,
                password: password
            })
        });
        
        const data = await response.json();
        
        if (response.ok && data.success) {
            // Login successful!
            console.log('Login successful');
            
            // Show success message briefly
            showSuccess('Login successful! Redirecting...');
            
            // Wait 1 second then redirect to dashboard
            setTimeout(function() {
                window.location.href = '/dashboard';
            }, 1000);
        } else {
            // Login failed
            showError(data.error || 'Login failed. Please try again.');
            resetButton();
        }
    } catch (error) {
        console.error('Login error:', error);
        showError('Connection error. Please try again.');
        resetButton();
    }
});

// Function to show error message
function showError(message) {
    errorMessage.innerHTML = message;
    errorMessage.style.display = 'block';
    errorMessage.className = 'error-message';
}

// Function to show success message
function showSuccess(message) {
    errorMessage.innerHTML = '✓ ' + message;
    errorMessage.style.display = 'block';
    errorMessage.className = 'error-message' + ' success';
    errorMessage.style.backgroundColor = '#d4edda';
    errorMessage.style.color = '#155724';
    errorMessage.style.borderColor = '#c3e6cb';
}

// Function to reset button after error
function resetButton() {
    submitBtn.disabled = false;
    submitBtn.classList.remove('loading');
    submitBtn.innerHTML = 'Login';
}

// Function to validate email format
function isValidEmail(email) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
}

// Allow Enter key to submit form
passwordInput.addEventListener('keypress', function(e) {
    if (e.key === 'Enter') {
        loginForm.dispatchEvent(new Event('submit'));
    }
});

// Clear error message when user starts typing
emailInput.addEventListener('focus', function() {
    if (errorMessage.style.display !== 'none') {
        errorMessage.style.display = 'none';
    }
});

passwordInput.addEventListener('focus', function() {
    if (errorMessage.style.display !== 'none') {
        errorMessage.style.display = 'none';
    }
});

// Add some interactivity - focus effects
emailInput.addEventListener('focus', function() {
    this.style.borderColor = '#667eea';
});

emailInput.addEventListener('blur', function() {
    this.style.borderColor = '#ddd';
});

passwordInput.addEventListener('focus', function() {
    this.style.borderColor = '#667eea';
});

passwordInput.addEventListener('blur', function() {
    this.style.borderColor = '#ddd';
});

console.log('Login form ready!');
