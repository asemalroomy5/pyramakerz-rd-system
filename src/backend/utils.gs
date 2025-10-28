/**
 * Enhanced password validation for manager access
 */
function validateEnhancedPassword(password) {
  const config = getConfig();
  Logger.log("Enhanced password validation for: " + password);
  
  // Check against all passwords in the PASSWORDS object
  const passwords = config.PASSWORDS;
  let userInfo = null;
  
  if (password === passwords.CTO) {
    userInfo = {
      name: "CTO",
      displayName: "CTO",
      department: "All",
      role: "CTO"
    };
  } else if (password === passwords.EMBEDDED_SENIOR) {
    userInfo = {
      name: "Embedded Senior",
      displayName: "Eng. Seif",
      department: "Embedded", 
      role: "Senior"
    };
  } else if (password === passwords.MECHANICAL_SENIOR) {
    userInfo = {
      name: "Mechanical Senior",
      displayName: "Eng. Mohaned",
      department: "Mechanical",
      role: "Senior"
    };
  } else if (password === passwords.AI_SENIOR) {
    userInfo = {
      name: "AI Senior", 
      displayName: "Eng. Maged",
      department: "AI",
      role: "Senior"
    };
  }
  
  if (userInfo) {
    Logger.log("✅ Password valid for: " + userInfo.displayName);
    return {
      isValid: true,
      userInfo: userInfo
    };
  } else {
    Logger.log("❌ Invalid password");
    return {
      isValid: false,
      userInfo: null
    };
  }
}

/**
 * Fallback to old password system for compatibility
 */
function validatePassword(password) {
  const config = getConfig();
  Logger.log("Legacy password validation for: " + password);
  
  // Check against old MANAGER_PASSWORD
  if (password === config.MANAGER_PASSWORD) {
    return {
      isValid: true,
      userInfo: {
        name: "Manager",
        displayName: "Manager", 
        department: "All",
        role: "Manager"
      }
    };
  }
  
  // Also check against new password system for backward compatibility
  return validateEnhancedPassword(password);
}

/**
 * Debug enhanced password validation
 */
function debugEnhancedPassword(inputPwd) {
  const config = getConfig();
  const result = validateEnhancedPassword(inputPwd);
  
  return {
    input: inputPwd,
    isValid: result.isValid,
    role: result.role,
    userInfo: result.userInfo,
    availableRoles: Object.keys(config.PASSWORDS)
  };
}

/**
 * Get user info by role
 */
function getUserInfoByRole(role) {
  const config = getConfig();
  return config.USER_ROLES[role] || null;
}
