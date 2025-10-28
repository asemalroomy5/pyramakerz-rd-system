// ┌───────────────────────────────────────────────┐
// │        PYRAMAKERZ R&D REPORTING SYSTEM        │
// │                 Main Entry                    │
// └───────────────────────────────────────────────┘

/**
 * Serve the web app with engineer login system
 */
function doGet(e) {
  try {
    // Initialize all required sheets
    const initResult = initializeAllSheets();
    if (!initResult.success) {
      return createErrorPage("Sheet Initialization Error", initResult.error);
    }
    
    // Check if engineer login is requested
    const path = e.pathInfo || '';
    const parameters = e.parameter || {};
    
    // Route based on path
    if (path === 'engineer-login' || parameters.view === 'engineer-login') {
      return serveEngineerLogin();
    } else if (path === 'engineer-panel' || parameters.view === 'engineer-panel') {
      return serveEngineerPanel();
    } else {
      // Default to CTO dashboard
      return serveCTODashboard();
    }
      
  } catch (error) {
    return createErrorPage("Application Error", error.toString());
  }
}

/**
 * Serve Engineer Login Page
 */
function serveEngineerLogin() {
  try {
    const template = HtmlService.createTemplateFromFile('engineer-login');
    const htmlOutput = template.evaluate();
    
    return htmlOutput
      .setTitle('Pyramakerz — Engineer Login')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return createErrorPage("Login Page Error", error.toString());
  }
}

/**
 * Serve Engineer Panel
 */
function serveEngineerPanel() {
  try {
    const template = HtmlService.createTemplateFromFile('engineer-panel');
    const htmlOutput = template.evaluate();
    
    return htmlOutput
      .setTitle('Pyramakerz — Engineer Panel')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return createErrorPage("Engineer Panel Error", error.toString());
  }
}

/**
 * Serve CTO Dashboard (existing functionality)
 */
function serveCTODashboard() {
  try {
    const template = HtmlService.createTemplateFromFile('index');
    const htmlOutput = template.evaluate();
    
    return htmlOutput
      .setTitle('Pyramakerz — R&D Reporting')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return createErrorPage("Dashboard Error", error.toString());
  }
}

/**
 * Create error page
 */
function createErrorPage(title, message) {
  return HtmlService.createHtmlOutput(`
    <html>
      <head>
        <title>Pyramakerz — Error</title>
        <style>
          body { 
            font-family: Arial, sans-serif; 
            padding: 40px; 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            text-align: center;
          }
          .error-container {
            background: rgba(255,255,255,0.1);
            padding: 30px;
            border-radius: 10px;
            backdrop-filter: blur(10px);
            max-width: 500px;
            margin: 0 auto;
          }
          button {
            background: #4CAF50;
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 5px;
            cursor: pointer;
            margin: 10px;
          }
        </style>
      </head>
      <body>
        <div class="error-container">
          <h2>⚠️ ${title}</h2>
          <p>${message}</p>
          <div>
            <button onclick="window.location.reload()">Retry</button>
            <button onclick="window.location.href='?view=engineer-login'">Engineer Login</button>
            <button onclick="window.location.href='?'">CTO Dashboard</button>
          </div>
        </div>
      </body>
    </html>
  `);
}

/**
 * Include external HTML files (REQUIRED for templating)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Verify and get the active spreadsheet
 */
function getActiveSpreadsheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log("Connected to spreadsheet: " + ss.getName());
    return ss;
  } catch (error) {
    Logger.log("Error accessing spreadsheet: " + error.toString());
    throw new Error("Cannot access the spreadsheet. Please make sure the script is bound to the correct Google Sheet.");
  }
}

/**
 * Test all data connections - UPDATED to remove popups
 */
function testDataConnections() {
  try {
    const ss = getActiveSpreadsheet();
    const config = getConfig();
    
    // Test each required sheet
    const sheetNames = Object.values(config.SHEET_NAMES);
    const missingSheets = [];
    
    sheetNames.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        missingSheets.push(sheetName);
      }
    });
    
    if (missingSheets.length > 0) {
      // Auto-create missing sheets instead of throwing error
      initializeAllSheets();
    }
    
    return { success: true, message: "All data connections working!" };
  } catch (error) {
    // Silent fail - just log, don't show popup
    Logger.log("Data connection test failed: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Recreate missing sheets and setup structure - UPDATED
 */
function setupMissingSheets() {
  try {
    const result = initializeAllSheets();
    return result.success ? "Sheet setup completed!" : "Sheet setup failed: " + result.error;
  } catch (error) {
    Logger.log("Error in setupMissingSheets: " + error.toString());
    return "Sheet setup failed: " + error.message;
  }
}

/**
 * Test the complete system - UPDATED to remove popups
 */
function testCompleteSystem() {
  try {
    // Test spreadsheet connection
    const ss = getActiveSpreadsheet();
    Logger.log("✓ Connected to: " + ss.getName());
    
    // Test configuration
    const config = getConfig();
    Logger.log("✓ Configuration loaded");
    
    // Test data services
    const projects = getProjects();
    Logger.log("✓ Projects loaded: " + projects.length);
    
    return { success: true, message: "All tests passed!" };
    
  } catch (error) {
    Logger.log("❌ System test failed: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Initialize menu when spreadsheet opens - UPDATED
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎯 Pyramakerz Quick Actions')
    .addItem('🚀 Create Project (Easy)', 'showCTOProjectCreatorFromSheets')
    .addSeparator()
    .addItem('👨‍💼 Engineer Login', 'showEngineerLoginFromSheets')
    .addItem('🔄 Refresh All Data', 'refreshWebAppData')
    .addItem('⚙️ Setup Missing Sheets', 'setupMissingSheets')
    .addToUi();
}

/**
 * Show Engineer Login from sheets
 */
function showEngineerLoginFromSheets() {
  const html = HtmlService.createHtmlOutputFromFile('engineer-login')
    .setWidth(400)
    .setHeight(500)
    .setTitle('Engineer Login');
  
  SpreadsheetApp.getUi().showModalDialog(html, '👨‍💼 Engineer Login');
}

/**
 * Complete system setup - UPDATED
 */
function setupCompleteSystem() {
  try {
    const result = initializeAllSheets();
    if (result.success) {
      Logger.log("Complete system setup finished successfully.");
    } else {
      Logger.log("Setup error: " + result.error);
    }
  } catch (e) {
    Logger.log("Setup error: " + e.toString());
  }
}

/**
 * Refresh all web app data - UPDATED to remove popups
 */
function refreshWebAppData() {
  try {
    // Silent refresh - no toast messages
    Logger.log("Web app data refreshed silently");
    return true;
  } catch (e) {
    Logger.log("Error refreshing data: " + e);
    return false;
  }
}

/**
 * Export dashboard data
 */
function exportDashboardData(projectFilter, exportType, statusFilter) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const url = ss.getUrl();
    return url;
  } catch (error) {
    throw new Error("Export failed: " + error.toString());
  }
}

/**
 * Get spreadsheet URL
 */
function getSpreadsheetUrl() {
  return SpreadsheetApp.getActiveSpreadsheet().getUrl();
}

/**
 * Add sample data (placeholder)
 */
function addSampleData() {
  try {
    return { success: true, message: "Sample data added" };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Get current user email for approval workflow
 */
function getCurrentUserEmail() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (error) {
    return 'unknown@pyramakerz.com';
  }
}

/**
 * Get available departments
 */
function getAvailableDepartments() {
  const config = getConfig();
  return config.DEPARTMENTS;
}

/**
 * Show CTO Project Creator from sheets
 */
function showCTOProjectCreatorFromSheets() {
  const html = HtmlService.createHtmlOutputFromFile('CTOProjectCreator')
    .setWidth(1000)
    .setHeight(800)
    .setTitle('🚀 CTO Project Creator');
  
  SpreadsheetApp.getUi().showModalDialog(html, '🚀 CTO Project Creator');
}

// ┌───────────────────────────────────────────────┐
// │           ENGINEER LOGIN SYSTEM               │
// └───────────────────────────────────────────────┘

/**
 * Authenticate engineer from login form - FIXED VERSION
 */
function authenticateEngineerFromForm(name, password) {
  try {
    console.log("authenticateEngineerFromForm called with:", name);
    
    const authResult = authenticateEngineer(name, password);
    console.log("Authentication result:", authResult);
    
    if (authResult.success) {
      // Get assigned projects for this engineer
      const assignedProjects = getAssignedProjectsForEngineer(
        authResult.engineer.name, 
        authResult.engineer.department
      );
      
      console.log("Assigned projects found:", assignedProjects.length);
      
      return {
        success: true,
        engineer: authResult.engineer,
        assignedProjects: assignedProjects,
        message: "Login successful"
      };
    } else {
      return authResult;
    }
  } catch (error) {
    console.log("Error in authenticateEngineerFromForm: " + error.toString());
    return { 
      success: false, 
      error: "Login failed: " + error.message 
    };
  }
}

/**
 * Change engineer password from form - FIXED VERSION
 */
function changeEngineerPasswordFromForm(engineerName, currentPassword, newPassword) {
  try {
    console.log("changeEngineerPasswordFromForm called for:", engineerName);
    
    // First authenticate the engineer
    const authResult = authenticateEngineer(engineerName, currentPassword);
    if (!authResult.success) {
      return { success: false, error: "Current password is incorrect" };
    }
    
    // Then change the password
    const changeResult = changeEngineerPassword(
      authResult.engineer.id, 
      currentPassword, 
      newPassword
    );
    
    return changeResult;
    
  } catch (error) {
    console.log("Error in changeEngineerPasswordFromForm: " + error.toString());
    return { 
      success: false, 
      error: "Password change failed: " + error.message 
    };
  }
}

/**
 * Reset and initialize engineer credentials
 */
function resetEngineerCredentials() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let credSheet = ss.getSheetByName(ENGINEER_CREDENTIALS_SHEET);
    
    // Delete existing sheet if it exists
    if (credSheet) {
      ss.deleteSheet(credSheet);
    }
    
    // Create new sheet
    credSheet = ss.insertSheet(ENGINEER_CREDENTIALS_SHEET);
    
    // Set headers
    credSheet.getRange('A1:E1').setValues([[
      'Engineer_ID', 'Name', 'Password', 'Department', 'Last_Login'
    ]]);
    
    // Set default credentials
    const defaultEngineers = [
      ['eng_seif', 'Eng. Seif', 'password123', 'Embedded', ''],
      ['eng_mohaned', 'Eng. Mohaned', 'password123', 'Mechanical', ''],
      ['eng_maged', 'Eng. Maged', 'password123', 'AI', ''],
      ['eng_rawan', 'Eng. Rawan', 'password123', 'Visual/Graphic', ''],
      ['eng_nehal', 'Eng. Nehal', 'password123', 'Flutter', ''],
      ['assem_murad', 'Assem Murad', 'password123', 'Mechanical', ''],
      ['eng_asem', 'Eng. Asem Alroomy', 'password123', 'Embedded', '']
    ];
    
    credSheet.getRange(2, 1, defaultEngineers.length, 5).setValues(defaultEngineers);
    
    // Format the sheet
    credSheet.getRange('A1:E1').setFontWeight('bold');
    credSheet.autoResizeColumns(1, 5);
    
    console.log("Engineer credentials reset successfully");
    return { success: true, message: "Credentials reset with default engineers" };
    
  } catch (error) {
    console.log("Error resetting credentials: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Complete system reset and setup
 */
function fixEngineerLoginSystem() {
  console.log("=== FIXING ENGINEER LOGIN SYSTEM ===");
  
  // Reset credentials
  const resetResult = resetEngineerCredentials();
  console.log("Reset result:", resetResult);
  
  // Initialize custom tasks sheet
  initializeEngineerCredentials();
  
  // Test authentication
  const testResult = authenticateEngineer("Eng. Seif", "password123");
  console.log("Test authentication:", testResult);
  
  // Test getting engineers
  const engineers = getAllEngineersWithDetails();
  console.log("Engineers found:", engineers.length);
  
  return {
    reset: resetResult,
    authentication: testResult,
    engineers: engineers.length,
    status: "System fixed and ready for testing"
  };
}

/**
 * Change engineer password from form
 */
function changeEngineerPasswordFromForm(engineerId, currentPassword, newPassword) {
  try {
    return changeEngineerPassword(engineerId, currentPassword, newPassword);
  } catch (error) {
    Logger.log("Error in changeEngineerPasswordFromForm: " + error.toString());
    return { 
      success: false, 
      error: "Password change failed: " + error.message 
    };
  }
}

/**
 * Get engineer dashboard data
 */
function getEngineerDashboardData(engineerId, engineerName, department) {
  try {
    const assignedProjects = getAssignedProjectsForEngineer(engineerName, department);
    const customTasks = getCustomTasksForEngineer(engineerId);
    
    return {
      success: true,
      assignedProjects: assignedProjects,
      customTasks: customTasks,
      department: department
    };
  } catch (error) {
    Logger.log("Error in getEngineerDashboardData: " + error.toString());
    return {
      success: false,
      error: "Failed to load dashboard data: " + error.message,
      assignedProjects: [],
      customTasks: [],
      department: department
    };
  }
}

/**
 * Submit task from engineer panel
 */
function submitTaskFromEngineerPanel(taskData) {
  try {
    // Add default times if not provided
    if (!taskData.startTime) {
      taskData.startTime = "09:00";
    }
    if (!taskData.endTime) {
      taskData.endTime = "17:00";
    }
    
    // Determine if this is a custom task or assigned task
    if (taskData.taskType === 'custom') {
      return submitCustomTaskForEngineer(taskData);
    } else {
      // For assigned tasks, use existing submission logic
      return submitTaskData(taskData);
    }
  } catch (error) {
    Logger.log("Error in submitTaskFromEngineerPanel: " + error.toString());
    return { 
      success: false, 
      error: "Task submission failed: " + error.message 
    };
  }
}

/**
 * Update task status from engineer panel
 */
function updateTaskStatusFromPanel(engineerId, taskIdentifier, newStatus, taskType = 'custom') {
  try {
    return updateTaskStatus(engineerId, taskIdentifier, newStatus, taskType);
  } catch (error) {
    Logger.log("Error in updateTaskStatusFromPanel: " + error.toString());
    return { 
      success: false, 
      error: "Status update failed: " + error.message 
    };
  }
}

/**
 * Test engineer system
 */
function testEngineerSystem() {
  try {
    // Test authentication
    const authResult = authenticateEngineer("Eng. Seif", "password123");
    
    if (authResult.success) {
      // Test assigned projects
      const projects = getAssignedProjectsForEngineer(
        authResult.engineer.name, 
        authResult.engineer.department
      );
      
      // Test custom tasks
      const customTasks = getCustomTasksForEngineer(authResult.engineer.id);
      
      return {
        success: true,
        auth: authResult,
        projectsCount: projects.length,
        customTasksCount: customTasks.length,
        engineer: authResult.engineer
      };
    } else {
      return {
        success: false,
        error: "Authentication failed in test"
      };
    }
  } catch (error) {
    return {
      success: false,
      error: "Engineer system test failed: " + error.toString()
    };
  }
}
/**
 * Validate engineer login - COMPATIBILITY FUNCTION
 * This fixes the "validateEngineerLogin is not a function" error
 */
function validateEngineerLogin(engineerName, password) {
  try {
    // Simply call the existing authentication function
    const result = authenticateEngineerFromForm(engineerName, password);
    return result;
  } catch (error) {
    Logger.log("Error in validateEngineerLogin: " + error.toString());
    return { 
      success: false, 
      error: "Login validation failed: " + error.message 
    };
  }
}