// --- INPUT SANITIZATION & XSS PROTECTION ---

/**
 * Sanitize HTML to prevent XSS attacks
 * Escapes dangerous characters that could be used for script injection
 * @param {string} input - User input to sanitize
 * @returns {string} Sanitized string safe for HTML display
 */
function sanitizeHtml(input) {
  if (input === null || input === undefined) {
    return '';
  }
  
  return String(input)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;')
    .replace(/\//g, '&#x2F;'); // Extra protection for closing tags
}

/**
 * Sanitize for use in JavaScript strings
 * @param {string} input - User input to sanitize
 * @returns {string} Sanitized string safe for JS
 */
function sanitizeJavaScript(input) {
  if (input === null || input === undefined) {
    return '';
  }
  
  return String(input)
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'")
    .replace(/"/g, '\\"')
    .replace(/\n/g, '\\n')
    .replace(/\r/g, '\\r')
    .replace(/\t/g, '\\t')
    .replace(/\//g, '\\/');
}

/**
 * Sanitize for use in HTML attributes
 * @param {string} input - User input to sanitize
 * @returns {string} Sanitized string safe for HTML attributes
 */
function sanitizeAttribute(input) {
  if (input === null || input === undefined) {
    return '';
  }
  
  return String(input)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#x27;')
    .replace(/\//g, '&#x2F;');
}

/**
 * Sanitize URL to prevent javascript: and data: URI attacks
 * @param {string} url - URL to sanitize
 * @returns {string} Safe URL or empty string if dangerous
 */
function sanitizeUrl(url) {
  if (!url) return '';
  
  const urlStr = String(url).trim().toLowerCase();
  
  // Block dangerous protocols
  const dangerousProtocols = ['javascript:', 'data:', 'vbscript:', 'file:'];
  for (let i = 0; i < dangerousProtocols.length; i++) {
    if (urlStr.indexOf(dangerousProtocols[i]) === 0) {
      Logger.log('Blocked dangerous URL: ' + url);
      return '';
    }
  }
  
  return String(url);
}

/**
 * Sanitize email address
 * @param {string} email - Email to sanitize
 * @returns {string} Sanitized email or empty if invalid
 */
function sanitizeEmail(email) {
  if (!email) return '';
  
  const emailStr = String(email).trim();
  
  // Basic email validation regex
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  
  if (!emailRegex.test(emailStr)) {
    return '';
  }
  
  // Remove any HTML tags
  return sanitizeHtml(emailStr);
}

/**
 * Sanitize student data object
 * @param {object} student - Student data to sanitize
 * @returns {object} Sanitized student data
 */
function sanitizeStudentData(student) {
  return {
    studentId: sanitizeHtml(student.studentId || ''),
    firstName: sanitizeHtml(student.firstName || ''),
    lastName: sanitizeHtml(student.lastName || ''),
    email: sanitizeEmail(student.email || ''),
    dob: student.dob || '', // Date object, safe
    gender: sanitizeHtml(student.gender || ''),
    enrollmentDate: student.enrollmentDate || '', // Date object, safe
    status: sanitizeHtml(student.status || ''),
    class: sanitizeHtml(student.class || '')
  };
}

/**
 * Sanitize user data object
 * @param {object} user - User data to sanitize
 * @returns {object} Sanitized user data
 */
function sanitizeUserData(user) {
  return {
    userId: sanitizeHtml(user.userId || ''),
    googleEmail: sanitizeEmail(user.googleEmail || ''),
    fullName: sanitizeHtml(user.fullName || ''),
    role: sanitizeHtml(user.role || ''),
    accountStatus: sanitizeHtml(user.accountStatus || ''),
    username: sanitizeHtml(user.username || ''),
    password: user.password || '' // Never display actual password
  };
}

/**
 * Sanitize array of objects (generic)
 * @param {array} dataArray - Array of data objects
 * @param {function} sanitizeFunc - Specific sanitization function for object type
 * @returns {array} Sanitized array
 */
function sanitizeDataArray(dataArray, sanitizeFunc) {
  if (!Array.isArray(dataArray)) {
    return [];
  }
  
  return dataArray.map(function(item) {
    return sanitizeFunc(item);
  });
}

/**
 * Remove potentially dangerous HTML tags and scripts
 * More aggressive than sanitizeHtml - strips all HTML
 * @param {string} input - Input possibly containing HTML
 * @returns {string} Plain text with HTML removed
 */
function stripHtml(input) {
  if (input === null || input === undefined) {
    return '';
  }
  
  return String(input)
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '') // Remove script tags
    .replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, '') // Remove style tags
    .replace(/<[^>]+>/g, '') // Remove all HTML tags
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#039;/g, "'")
    .trim();
}

/**
 * Validate and sanitize input based on type
 * @param {any} input - Input value
 * @param {string} type - Expected type (text, email, number, date, url)
 * @returns {any} Validated and sanitized value
 */
function validateAndSanitize(input, type) {
  switch(type) {
    case 'email':
      return sanitizeEmail(input);
    
    case 'url':
      return sanitizeUrl(input);
    
    case 'number':
      const num = Number(input);
      return isNaN(num) ? 0 : num;
    
    case 'date':
      try {
        return new Date(input);
      } catch (e) {
        return '';
      }
    
    case 'text':
    default:
      return sanitizeHtml(input);
  }
}

/**
 * Sanitize SQL-like input (for sheet queries)
 * Prevents injection attacks in sheet queries
 * @param {string} input - Input to sanitize
 * @returns {string} Sanitized input
 */
function sanitizeSheetQuery(input) {
  if (input === null || input === undefined) {
    return '';
  }
  
  return String(input)
    .replace(/'/g, "''") // Escape single quotes
    .replace(/;/g, '') // Remove semicolons
    .replace(/--/g, '') // Remove SQL comments
    .replace(/\/\*/g, '') // Remove multi-line comment start
    .replace(/\*\//g, ''); // Remove multi-line comment end
}

/**
 * Create Content Security Policy header value
 * @returns {string} CSP header value
 */
function getCSPHeader() {
  return "default-src 'self'; " +
         "script-src 'self' 'unsafe-inline' https://apis.google.com; " +
         "style-src 'self' 'unsafe-inline'; " +
         "img-src 'self' data: https:; " +
         "font-src 'self' data:; " +
         "connect-src 'self' https://script.google.com; " +
         "frame-ancestors 'self';";
}

/**
 * Test sanitization functions
 * @returns {object} Test results
 */
function testSanitization() {
  const tests = [
    {
      name: 'XSS Script Tag',
      input: '<script>alert("XSS")</script>Hello',
      expected: '&lt;script&gt;alert(&quot;XSS&quot;)&lt;&#x2F;script&gt;Hello',
      result: sanitizeHtml('<script>alert("XSS")</script>Hello')
    },
    {
      name: 'XSS Image OnError',
      input: '<img src=x onerror="alert(1)">',
      expected: '&lt;img src=x onerror=&quot;alert(1)&quot;&gt;',
      result: sanitizeHtml('<img src=x onerror="alert(1)">')
    },
    {
      name: 'JavaScript URL',
      input: 'javascript:alert(1)',
      expected: '',
      result: sanitizeUrl('javascript:alert(1)')
    },
    {
      name: 'Valid Email',
      input: 'user@example.com',
      expected: 'user@example.com',
      result: sanitizeEmail('user@example.com')
    },
    {
      name: 'Invalid Email with Script',
      input: 'user@example.com<script>alert(1)</script>',
      expected: '',
      result: sanitizeEmail('user@example.com<script>alert(1)</script>')
    }
  ];
  
  let passed = 0;
  let failed = 0;
  
  tests.forEach(function(test) {
    if (test.result === test.expected) {
      passed++;
      Logger.log('✅ PASS: ' + test.name);
    } else {
      failed++;
      Logger.log('❌ FAIL: ' + test.name);
      Logger.log('   Expected: ' + test.expected);
      Logger.log('   Got: ' + test.result);
    }
  });
  
  return {
    total: tests.length,
    passed: passed,
    failed: failed,
    allPassed: failed === 0
  };
}
