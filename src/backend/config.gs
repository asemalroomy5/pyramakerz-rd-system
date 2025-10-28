// ┌───────────────────────────────────────────────┐
// │        PYRAMAKERZ R&D REPORTING SYSTEM        │
// │               Configuration                   │
// └───────────────────────────────────────────────┘

const CONFIG = {
  // Email Configuration
  CTO_EMAIL: "nemr@pyramakerz.com",
  SEIF_EMAIL: "Seif_El-Dein@pyramakerz.com",
  
  // Sheet Names - UPDATED with new sheets
  SHEET_NAMES: {
    EMBEDDED_RESPONSES: "Embedded_Responses",
    AI_RESPONSES: "AI_Data_Responses", 
    MECHANICAL_RESPONSES: "Mechanical_Responses",
    TEAM_MAPPING: "Team_Mapping",
    PROJECTS_TASKS: "Projects_Tasks",
    ACCESS_CONTROL: "Access_Control",
    ENGINEER_CREDENTIALS: "Engineer_Credentials",
    CUSTOM_TASKS: "Custom_Tasks"
  },
  
  // System Settings
  MAX_TASK_BLOCKS: 4,
  
  // Enhanced Password System for CTO and Seniors
  PASSWORDS: {
    CTO: "CTO2025",
    EMBEDDED_SENIOR: "EMB2025",
    MECHANICAL_SENIOR: "MEC2025", 
    AI_SENIOR: "AI2025"
  },
  
  // User Roles Mapping
  USER_ROLES: {
    CTO: {
      name: "CTO",
      displayName: "Chief Technology Officer",
      department: "All",
      email: "nemr@pyramakerz.com"
    },
    EMBEDDED_SENIOR: {
      name: "Eng. Seif",
      displayName: "Senior Embedded Engineer",
      department: "Embedded", 
      email: "eng.seif@pyramakerz.com"
    },
    MECHANICAL_SENIOR: {
      name: "Eng. Mohaned",
      displayName: "Senior Mechanical Engineer",
      department: "Mechanical",
      email: "eng.mohaned@pyramakerz.com"
    },
    AI_SENIOR: {
      name: "Eng. Maged", 
      displayName: "Senior AI Engineer",
      department: "AI",
      email: "eng.maged@pyramakerz.com"
    }
  },

  SENIOR_SHEET_SUFFIX: "_Team",
  
  // Departments
  DEPARTMENTS: ["Embedded", "AI", "Mechanical", "Visual/Graphic", "Flutter"],
  
  // Default Engineers - UPDATED with consistent naming
  DEFAULT_ENGINEERS: [
    { id: 'engmohaned', name: 'Eng. Mohaned', role: 'Senior', department: 'Mechanical', email: 'eng.mohaned@pyramakerz.com', level: 'Senior' },
    { id: 'assemmurad', name: 'Assem Murad', role: 'Engineer', department: 'Mechanical', email: 'assem.murad@pyramakerz.com', level: 'Junior' },
    { id: 'engseif', name: 'Eng. Seif', role: 'Senior', department: 'Embedded', email: 'eng.seif@pyramakerz.com', level: 'Senior' },
    { id: 'engasemalroomy', name: 'Eng. Asem Alroomy', role: 'Engineer', department: 'Embedded', email: 'eng.asem@pyramakerz.com', level: 'Junior' },
    { id: 'engrawan', name: 'Eng. Rawan', role: 'Senior', department: 'Visual/Graphic', email: 'eng.rawan@pyramakerz.com', level: 'Senior' },
    { id: 'engnehal', name: 'Eng. Nehal', role: 'Senior', department: 'Flutter', email: 'eng.nehal@pyramakerz.com', level: 'Senior' },
    { id: 'engmaged', name: 'Eng. Maged', role: 'Senior', department: 'AI', email: 'eng.maged@pyramakerz.com', level: 'Senior' }
  ],

  // Engineer Login System - NEW
  ENGINEER_LOGIN: {
    DEFAULT_PASSWORD: "password123",
    SESSION_TIMEOUT: 24 * 60 * 60 * 1000, // 24 hours in milliseconds
    PASSWORD_MIN_LENGTH: 6
  },

  // Task Status Options - NEW (Trello-like status system)
  TASK_STATUSES: [
    "Not Started",
    "In Progress", 
    "Blocked",
    "In Review",
    "Completed",
    "Done"
  ],

  // Default Work Hours - NEW (Auto Time Adjustment)
  DEFAULT_WORK_HOURS: {
    START: "09:00",
    END: "17:00"
  }
};

// Make config available globally
function getConfig() {
  return CONFIG;
}

/**
 * Initialize all required sheets for the system
 */
function initializeAllSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();
    
    // Initialize Engineer Credentials sheet
    let credSheet = ss.getSheetByName(config.SHEET_NAMES.ENGINEER_CREDENTIALS);
    if (!credSheet) {
      credSheet = ss.insertSheet(config.SHEET_NAMES.ENGINEER_CREDENTIALS);
      credSheet.getRange('A1:E1').setValues([[
        'Engineer_ID', 'Name', 'Password', 'Department', 'Last_Login'
      ]]);
      
      // Create default credentials based on DEFAULT_ENGINEERS
      const defaultCredentials = config.DEFAULT_ENGINEERS.map(engineer => [
        engineer.id,
        engineer.name,
        config.ENGINEER_LOGIN.DEFAULT_PASSWORD,
        engineer.department,
        '' // Last_Login empty initially
      ]);
      
      if (credSheet.getLastRow() === 1 && defaultCredentials.length > 0) {
        credSheet.getRange(2, 1, defaultCredentials.length, 5).setValues(defaultCredentials);
      }
    }
    
    // Initialize Custom Tasks sheet
    let customTasksSheet = ss.getSheetByName(config.SHEET_NAMES.CUSTOM_TASKS);
    if (!customTasksSheet) {
      customTasksSheet = ss.insertSheet(config.SHEET_NAMES.CUSTOM_TASKS);
      customTasksSheet.getRange('A1:J1').setValues([[
        'Timestamp', 'Engineer_ID', 'Engineer_Name', 'Project_Name', 'Task_Name',
        'Description', 'Start_Time', 'End_Time', 'Status', 'Checklist_Items'
      ]]);
      
      // Format the headers
      customTasksSheet.getRange('A1:J1').setFontWeight('bold');
    }
    
    // Ensure other required sheets exist
    const requiredSheets = [
      config.SHEET_NAMES.PROJECTS_TASKS,
      config.SHEET_NAMES.TEAM_MAPPING,
      config.SHEET_NAMES.EMBEDDED_RESPONSES,
      config.SHEET_NAMES.AI_RESPONSES,
      config.SHEET_NAMES.MECHANICAL_RESPONSES
    ];
    
    requiredSheets.forEach(sheetName => {
      if (!ss.getSheetByName(sheetName)) {
        ss.insertSheet(sheetName);
        Logger.log("Created missing sheet: " + sheetName);
      }
    });
    
    return {
      success: true,
      message: "All sheets initialized successfully",
      sheets: {
        credentials: config.SHEET_NAMES.ENGINEER_CREDENTIALS,
        customTasks: config.SHEET_NAMES.CUSTOM_TASKS
      }
    };
    
  } catch (error) {
    Logger.log("Error in initializeAllSheets: " + error.toString());
    return {
      success: false,
      error: "Failed to initialize sheets: " + error.message
    };
  }
}

/**
 * Get available status options for tasks
 */
function getTaskStatusOptions() {
  return getConfig().TASK_STATUSES;
}

/**
 * Get default work hours
 */
function getDefaultWorkHours() {
  return getConfig().DEFAULT_WORK_HOURS;
}

/**
 * Test configuration and sheet setup
 */
function testConfiguration() {
  try {
    const config = getConfig();
    const sheetInit = initializeAllSheets();
    
    return {
      config: {
        sheets: Object.keys(config.SHEET_NAMES).length,
        departments: config.DEPARTMENTS.length,
        engineers: config.DEFAULT_ENGINEERS.length,
        statusOptions: config.TASK_STATUSES.length
      },
      sheetInitialization: sheetInit,
      defaultWorkHours: config.DEFAULT_WORK_HOURS
    };
    
  } catch (error) {
    return {
      error: error.toString(),
      configTest: "FAILED"
    };
  }
}