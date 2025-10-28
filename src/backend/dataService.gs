// ┌───────────────────────────────────────────────┐
// │        PYRAMAKERZ R&D REPORTING SYSTEM        │
// │                Data Service                   │
// └───────────────────────────────────────────────┘

// Add to top of file - Engineer Authentication System
var ENGINEER_CREDENTIALS_SHEET = "Engineer_Credentials";
var CUSTOM_TASKS_SHEET = "Custom_Tasks";

/**
 * Initialize engineer credentials sheet if it doesn't exist
 */
function initializeEngineerCredentials() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let credSheet = ss.getSheetByName(ENGINEER_CREDENTIALS_SHEET);
    
    if (!credSheet) {
      credSheet = ss.insertSheet(ENGINEER_CREDENTIALS_SHEET);
      credSheet.getRange('A1:E1').setValues([[
        'Engineer_ID', 'Name', 'Password', 'Department', 'Last_Login'
      ]]);
      
      // Set default credentials for testing
      const defaultEngineers = [
        ['eng_seif', 'Eng. Seif', 'password123', 'Embedded', ''],
        ['eng_mohanad', 'Eng. Mohanad', 'password123', 'Mechanical', ''],
        ['eng_maged', 'Eng. Maged', 'password123', 'AI', ''],
        ['eng_ahmed', 'Eng. Ahmed', 'password123', 'Visual/Graphic', ''],
        ['eng_omar', 'Eng. Omar', 'password123', 'Flutter', '']
      ];
      
      if (credSheet.getLastRow() === 1) {
        credSheet.getRange(2, 1, defaultEngineers.length, 5).setValues(defaultEngineers);
      }
    }
    
    // Initialize Custom Tasks sheet
    let customTasksSheet = ss.getSheetByName(CUSTOM_TASKS_SHEET);
    if (!customTasksSheet) {
      customTasksSheet = ss.insertSheet(CUSTOM_TASKS_SHEET);
      customTasksSheet.getRange('A1:J1').setValues([[
        'Timestamp', 'Engineer_ID', 'Engineer_Name', 'Project_Name', 'Task_Name',
        'Description', 'Start_Time', 'End_Time', 'Status', 'Checklist_Items'
      ]]);
    }
    
    return true;
  } catch (error) {
    Logger.log("Error initializing engineer credentials: " + error.toString());
    return false;
  }
}

/**
 * Authenticate engineer with name and password
 */
function authenticateEngineer(name, password) {
  try {
    initializeEngineerCredentials();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const credSheet = ss.getSheetByName(ENGINEER_CREDENTIALS_SHEET);
    
    if (!credSheet) {
      return { success: false, error: "Credentials sheet not found" };
    }
    
    const data = credSheet.getDataRange().getValues();
    if (data.length < 2) {
      return { success: false, error: "No engineer credentials found" };
    }
    
    const headers = data[0];
    const nameIndex = headers.indexOf('Name');
    const passwordIndex = headers.indexOf('Password');
    const idIndex = headers.indexOf('Engineer_ID');
    const deptIndex = headers.indexOf('Department');
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const engineerName = String(row[nameIndex] || "").trim();
      const engineerPassword = String(row[passwordIndex] || "").trim();
      const engineerId = String(row[idIndex] || "").trim();
      const department = String(row[deptIndex] || "").trim();
      
      if (engineerName.toLowerCase() === name.toLowerCase() && 
          engineerPassword === password) {
        
        // Update last login
        credSheet.getRange(i + 1, headers.indexOf('Last_Login') + 1)
          .setValue(new Date());
        
        return {
          success: true,
          engineer: {
            id: engineerId,
            name: engineerName,
            department: department
          }
        };
      }
    }
    
    return { success: false, error: "Invalid credentials" };
    
  } catch (error) {
    Logger.log("Error in authenticateEngineer: " + error.toString());
    return { success: false, error: "Authentication failed: " + error.message };
  }
}

/**
 * Change engineer password
 */
function changeEngineerPassword(engineerId, currentPassword, newPassword) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const credSheet = ss.getSheetByName(ENGINEER_CREDENTIALS_SHEET);
    
    if (!credSheet) {
      return { success: false, error: "Credentials sheet not found" };
    }
    
    const data = credSheet.getDataRange().getValues();
    const headers = data[0];
    const idIndex = headers.indexOf('Engineer_ID');
    const passwordIndex = headers.indexOf('Password');
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowEngineerId = String(row[idIndex] || "").trim();
      const currentStoredPassword = String(row[passwordIndex] || "").trim();
      
      if (rowEngineerId === engineerId) {
        if (currentStoredPassword === currentPassword) {
          // Update password
          credSheet.getRange(i + 1, passwordIndex + 1).setValue(newPassword);
          return { success: true, message: "Password updated successfully" };
        } else {
          return { success: false, error: "Current password is incorrect" };
        }
      }
    }
    
    return { success: false, error: "Engineer not found" };
    
  } catch (error) {
    Logger.log("Error in changeEngineerPassword: " + error.toString());
    return { success: false, error: "Password change failed: " + error.message };
  }
}

/**
 * Get projects assigned to specific engineer - UPDATED for login system
 */
function getAssignedProjectsForEngineer(engineerName, engineerDepartment) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();
    const sh = ss.getSheetByName(config.SHEET_NAMES.PROJECTS_TASKS);
    
    if (!sh) {
      Logger.log("Projects_Tasks sheet not found");
      return getSampleProjects().filter(proj => 
        proj.departments.toLowerCase().includes(engineerDepartment.toLowerCase())
      );
    }
    
    const data = sh.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log("No data in Projects_Tasks sheet");
      return getSampleProjects().filter(proj => 
        proj.departments.toLowerCase().includes(engineerDepartment.toLowerCase())
      );
    }
    
    const assignedProjects = new Map();
    
    // First, find all tasks assigned to this engineer
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length < 11) continue;
      
      const rowType = String(row[0] || "").trim();
      if (rowType !== "Task") continue;
      
      const projectName = String(row[2] || "").trim();
      const assignedTo = String(row[8] || "").trim();
      const taskDepartment = String(row[9] || "").trim();
      
      // Check if task is assigned to this engineer or in their department
      const isAssigned = assignedTo.toLowerCase().includes(engineerName.toLowerCase()) ||
                        engineerName.toLowerCase().includes(assignedTo.toLowerCase());
      
      const isInDepartment = taskDepartment.toLowerCase() === engineerDepartment.toLowerCase();
      
      if (isAssigned || isInDepartment) {
        if (!assignedProjects.has(projectName)) {
          // Find project details
          for (let j = 1; j < data.length; j++) {
            const projRow = data[j];
            if (!projRow || projRow.length < 6) continue;
            
            const projType = String(projRow[0] || "").trim();
            const projName = String(projRow[2] || "").trim();
            
            if (projType === "Project" && projName === projectName) {
              assignedProjects.set(projectName, {
                id: String(projRow[1] || "").trim() || projectName.replace(/\s+/g, '').substring(0, 3).toUpperCase(),
                name: projectName,
                departments: String(projRow[3] || "").trim(),
                status: String(projRow[4] || "").trim(),
                manager: String(projRow[5] || "").trim()
              });
              break;
            }
          }
        }
      }
    }
    
    const projects = Array.from(assignedProjects.values());
    
    // If no assigned projects found, return department projects as fallback
    if (projects.length === 0) {
      return getProjects().filter(proj => 
        proj.departments.toLowerCase().includes(engineerDepartment.toLowerCase())
      );
    }
    
    Logger.log(`Found ${projects.length} assigned projects for ${engineerName}`);
    return projects;
    
  } catch (error) {
    Logger.log("Error in getAssignedProjectsForEngineer: " + error.toString());
    return getSampleProjects().filter(proj => 
      proj.departments.toLowerCase().includes(engineerDepartment.toLowerCase())
    );
  }
}

/**
 * Get tasks for project filtered by engineer assignment - UPDATED
 */
function getTasksForProjectForEngineer(projectName, engineerName, engineerDepartment) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();
    const sh = ss.getSheetByName(config.SHEET_NAMES.PROJECTS_TASKS);
    
    if (!sh) {
      throw new Error("Projects_Tasks sheet not found");
    }
    
    const data = sh.getDataRange().getValues();
    if (data.length < 2) {
      throw new Error("No data in Projects_Tasks sheet");
    }
    
    const tasks = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length < 11) continue;
      
      const rowType = String(row[0] || "").trim();
      if (rowType !== "Task") continue;
      
      const taskProjectName = String(row[2] || "").trim();
      const taskId = String(row[6] || "").trim();
      const taskName = String(row[7] || "").trim();
      const assignedTo = String(row[8] || "").trim();
      const department = String(row[9] || "").trim();
      const description = String(row[10] || "").trim();
      
      // Match project name and check assignment
      if (taskProjectName.toLowerCase() === projectName.toLowerCase()) {
        const isAssigned = assignedTo.toLowerCase().includes(engineerName.toLowerCase()) ||
                          engineerName.toLowerCase().includes(assignedTo.toLowerCase());
        
        const isInDepartment = department.toLowerCase() === engineerDepartment.toLowerCase();
        
        if (isAssigned || isInDepartment) {
          tasks.push({
            id: taskId || "TASK_" + (tasks.length + 1),
            name: taskName,
            description: description || taskName,
            assignedTo: assignedTo,
            department: department,
            start: "09:00", // Default to 9:00 AM
            end: "17:00",   // Default to 5:00 PM
            isAssigned: isAssigned
          });
        }
      }
    }
    
    if (tasks.length === 0) {
      Logger.log("No assigned tasks found for project: " + projectName);
    }
    
    return tasks;
    
  } catch (error) {
    Logger.log("Error in getTasksForProjectForEngineer: " + error.toString());
    throw new Error("Could not load tasks: " + error.message);
  }
}

/**
 * Get custom tasks for engineer with hide completed functionality
 */
function getCustomTasksForEngineer(engineerId, hideCompleted = true) {
  try {
    initializeEngineerCredentials();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const customSheet = ss.getSheetByName(CUSTOM_TASKS_SHEET);
    
    if (!customSheet) {
      return [];
    }
    
    const data = customSheet.getDataRange().getValues();
    if (data.length < 2) {
      return [];
    }
    
    const headers = data[0];
    const tasks = [];
    const today = new Date();
    const yesterday = new Date(today);
    yesterday.setDate(yesterday.getDate() - 1);
    const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const yesterdayStr = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length === 0) continue;
      
      const rowEngineerId = String(row[1] || "").trim();
      const timestamp = row[0];
      const status = String(row[8] || "").trim();
      
      // Filter by engineer and apply hide completed logic
      if (rowEngineerId === engineerId) {
        const taskDate = timestamp instanceof Date ? 
          Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy-MM-dd") : 
          todayStr;
        
        // Hide completed tasks from previous days
        if (hideCompleted && status.toLowerCase() === 'done' && taskDate !== todayStr && taskDate !== yesterdayStr) {
          continue;
        }
        
        tasks.push({
          timestamp: timestamp,
          engineerId: rowEngineerId,
          engineerName: String(row[2] || "").trim(),
          projectName: String(row[3] || "").trim(),
          taskName: String(row[4] || "").trim(),
          description: String(row[5] || "").trim(),
          startTime: String(row[6] || "09:00").trim(),
          endTime: String(row[7] || "17:00").trim(),
          status: status,
          checklistItems: String(row[9] || "").trim(),
          taskDate: taskDate
        });
      }
    }
    
    // Sort by timestamp (newest first)
    tasks.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    
    return tasks;
    
  } catch (error) {
    Logger.log("Error in getCustomTasksForEngineer: " + error.toString());
    return [];
  }
}

/**
 * Submit custom task for engineer
 */
function submitCustomTaskForEngineer(taskData) {
  try {
    initializeEngineerCredentials();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const customSheet = ss.getSheetByName(CUSTOM_TASKS_SHEET);
    
    if (!customSheet) {
      return { success: false, error: "Custom tasks sheet not found" };
    }
    
    const timestamp = new Date();
    const row = [
      timestamp,
      taskData.engineerId,
      taskData.engineerName,
      taskData.projectName || "Custom Project",
      taskData.taskName,
      taskData.description,
      taskData.startTime || "09:00",
      taskData.endTime || "17:00",
      taskData.status || "In Progress",
      taskData.checklistItems || ""
    ];
    
    customSheet.appendRow(row);
    
    return { success: true, message: "Custom task submitted successfully" };
    
  } catch (error) {
    Logger.log("Error in submitCustomTaskForEngineer: " + error.toString());
    return { success: false, error: "Failed to submit custom task: " + error.message };
  }
}

/**
 * Update task status (Trello-like status system)
 */
function updateTaskStatus(engineerId, taskIdentifier, newStatus, taskType = 'custom') {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (taskType === 'custom') {
      const customSheet = ss.getSheetByName(CUSTOM_TASKS_SHEET);
      if (!customSheet) return { success: false, error: "Custom tasks sheet not found" };
      
      const data = customSheet.getDataRange().getValues();
      const headers = data[0];
      const statusIndex = headers.indexOf('Status');
      const engineerIdIndex = headers.indexOf('Engineer_ID');
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowEngineerId = String(row[engineerIdIndex] || "").trim();
        const rowTaskName = String(row[4] || "").trim(); // Task_Name column
        
        if (rowEngineerId === engineerId && rowTaskName === taskIdentifier) {
          customSheet.getRange(i + 1, statusIndex + 1).setValue(newStatus);
          return { success: true, message: "Status updated successfully" };
        }
      }
    }
    
    return { success: false, error: "Task not found" };
    
  } catch (error) {
    Logger.log("Error in updateTaskStatus: " + error.toString());
    return { success: false, error: "Failed to update status: " + error.message };
  }
}

/**
 * Get completed tasks for date (for hide completed functionality)
 */
function getCompletedTasksForDate(engineerId, dateStr) {
  try {
    const customTasks = getCustomTasksForEngineer(engineerId, false);
    return customTasks.filter(task => 
      task.status.toLowerCase() === 'done' && 
      task.taskDate === dateStr
    );
  } catch (error) {
    Logger.log("Error in getCompletedTasksForDate: " + error.toString());
    return [];
  }
}

// ┌───────────────────────────────────────────────┐
// │          EXISTING FUNCTIONS (UPDATED)         │
// └───────────────────────────────────────────────┘

/**
 * Get projects from Projects_Tasks sheet - UPDATED with engineer filter
 */
function getProjects(department = null, engineerName = null) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();
    
    if (!ss) {
      return getSampleProjects(department);
    }
    
    const sh = ss.getSheetByName(config.SHEET_NAMES.PROJECTS_TASKS);
    
    if (!sh) {
      Logger.log("Projects_Tasks sheet not found");
      return getSampleProjects(department);
    }
    
    const data = sh.getDataRange().getValues();
    if (data.length < 2) {
      return getSampleProjects(department);
    }
    
    const projects = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length === 0) continue;
      
      const rowType = String(row[0] || "").trim();
      if (rowType !== "Project") continue;
      
      const projID = String(row[1] || "").trim();
      const projName = String(row[2] || "").trim();
      const depts = String(row[3] || "").trim();
      const status = String(row[4] || "").trim();
      const manager = String(row[5] || "").trim();
      
      if (!projName) continue;
      
      // Apply department filter
      if (department && !depts.toLowerCase().includes(department.toLowerCase())) {
        continue;
      }
      
      // Apply engineer filter if provided
      if (engineerName) {
        const hasAssignedTasks = checkProjectHasEngineerTasks(projName, engineerName, department);
        if (!hasAssignedTasks) continue;
      }
      
      projects.push({ 
        id: projID || projName.replace(/\s+/g, '').substring(0, 3).toUpperCase(), 
        name: projName, 
        departments: depts, 
        status: status, 
        manager: manager 
      });
    }
    
    if (projects.length === 0) {
      return getSampleProjects(department);
    }
    
    Logger.log("Found " + projects.length + " projects");
    return projects;
  } catch (error) {
    Logger.log("Error in getProjects: " + error.toString());
    return getSampleProjects(department);
  }
}

/**
 * Helper function to check if project has tasks for specific engineer
 */
function checkProjectHasEngineerTasks(projectName, engineerName, department) {
  try {
    const tasks = getTasksForProjectForEngineer(projectName, engineerName, department);
    return tasks.length > 0;
  } catch (error) {
    return false;
  }
}

/**
 * Get tasks for a project from Projects_Tasks sheet - UPDATED
 */
function getTasksForProject(projectName, engineerName = null, engineerDepartment = null) {
  try {
    // If engineer info provided, use filtered version
    if (engineerName && engineerDepartment) {
      return getTasksForProjectForEngineer(projectName, engineerName, engineerDepartment);
    }
    
    // Otherwise use original logic
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();
    const sh = ss.getSheetByName(config.SHEET_NAMES.PROJECTS_TASKS);
    
    if (!sh) {
      throw new Error("Projects_Tasks sheet not found");
    }
    
    const data = sh.getDataRange().getValues();
    if (data.length < 2) {
      throw new Error("No data in Projects_Tasks sheet");
    }
    
    const tasks = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length < 11) continue;
      
      const rowType = String(row[0] || "").trim();
      if (rowType !== "Task") continue;
      
      const taskProjectName = String(row[2] || "").trim();
      const taskId = String(row[6] || "").trim();
      const taskName = String(row[7] || "").trim();
      const assignedTo = String(row[8] || "").trim();
      const department = String(row[9] || "").trim();
      const description = String(row[10] || "").trim();
      
      if (taskProjectName.toLowerCase() !== projectName.toLowerCase()) continue;
      
      if (taskName) {
        tasks.push({
          id: taskId || "TASK_" + (tasks.length + 1),
          name: taskName,
          description: description || taskName,
          assignedTo: assignedTo,
          department: department,
          start: "09:00",
          end: "17:00"
        });
      }
    }
    
    if (tasks.length === 0) {
      Logger.log("No tasks found for project: " + projectName);
    }
    
    return tasks;
    
  } catch (error) {
    Logger.log("Error in getTasksForProject: " + error.toString());
    throw new Error("Could not load tasks from sheet: " + error.message);
  }
}

/**
 * Get tasks assigned to specific engineer - UPDATED with date filtering
 */
function getTasksForEngineer(engineerName, department = null, hideCompleted = true) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();
    const sh = ss.getSheetByName(config.SHEET_NAMES.PROJECTS_TASKS);
    
    if (!sh) {
      throw new Error("Projects_Tasks sheet not found");
    }
    
    const data = sh.getDataRange().getValues();
    if (data.length < 2) {
      throw new Error("No data in Projects_Tasks sheet");
    }
    
    const tasks = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length < 11) continue;
      
      const rowType = String(row[0] || "").trim();
      if (rowType !== "Task") continue;
      
      const taskProjectName = String(row[2] || "").trim();
      const taskId = String(row[6] || "").trim();
      const taskName = String(row[7] || "").trim();
      const assignedTo = String(row[8] || "").trim();
      const taskDepartment = String(row[9] || "").trim();
      const description = String(row[10] || "").trim();
      
      if (assignedTo.toLowerCase().includes(engineerName.toLowerCase()) ||
          engineerName.toLowerCase().includes(assignedTo.toLowerCase())) {
        
        if (department && taskDepartment.toLowerCase() !== department.toLowerCase()) {
          continue;
        }
        
        tasks.push({
          id: taskId || "TASK_" + (tasks.length + 1),
          name: taskName,
          description: description || taskName,
          assignedTo: assignedTo,
          department: taskDepartment,
          projectName: taskProjectName,
          start: "09:00",
          end: "17:00"
        });
      }
    }
    
    Logger.log(`Found ${tasks.length} assigned tasks for ${engineerName}`);
    return tasks;
    
  } catch (error) {
    Logger.log("Error in getTasksForEngineer: " + error.toString());
    throw new Error("Could not load assigned tasks: " + error.message);
  }
}

// Keep all other existing functions unchanged...
// [All your existing functions remain here without changes]
// Only the functions above were modified or added

/**
 * Test the new engineer authentication system
 */
function testEngineerAuthSystem() {
  try {
    // Initialize sheets
    initializeEngineerCredentials();
    
    // Test authentication
    const authResult = authenticateEngineer("Eng. Seif", "password123");
    Logger.log("Auth Test Result: " + JSON.stringify(authResult));
    
    if (authResult.success) {
      // Test assigned projects
      const projects = getAssignedProjectsForEngineer(
        authResult.engineer.name, 
        authResult.engineer.department
      );
      Logger.log("Assigned Projects: " + projects.length);
      
      // Test custom tasks
      const customTasks = getCustomTasksForEngineer(authResult.engineer.id);
      Logger.log("Custom Tasks: " + customTasks.length);
    }
    
    return {
      auth: authResult,
      sheetsInitialized: true
    };
    
  } catch (error) {
    Logger.log("Test Error: " + error.toString());
    return { error: error.toString() };
  }
}
/**
 * Get all tasks for engineer (assigned + department projects)
 */
function getAllTasksForEngineer(engineerName, department) {
  try {
    const assignedTasks = getTasksForEngineer(engineerName, department);
    const departmentTasks = getTasksForDepartment(department);
    
    // Combine and remove duplicates
    const allTasks = [...assignedTasks];
    const assignedTaskNames = new Set(assignedTasks.map(t => t.name));
    
    departmentTasks.forEach(task => {
      if (!assignedTaskNames.has(task.name)) {
        allTasks.push(task);
      }
    });
    
    return allTasks;
    
  } catch (error) {
    Logger.log("Error in getAllTasksForEngineer: " + error.toString());
    return getTasksForEngineer(engineerName, department);
  }
}

/**
 * Get tasks for a department (fallback if no assigned tasks)
 */
function getTasksForDepartment(department) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();
    const sh = ss.getSheetByName(config.SHEET_NAMES.PROJECTS_TASKS);
    
    if (!sh) return [];
    
    const data = sh.getDataRange().getValues();
    const tasks = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length < 11) continue;
      
      const rowType = String(row[0] || "").trim();
      if (rowType !== "Task") continue;
      
      const taskDepartment = String(row[9] || "").trim();
      const taskProjectName = String(row[2] || "").trim();
      const taskId = String(row[6] || "").trim();
      const taskName = String(row[7] || "").trim();
      const assignedTo = String(row[8] || "").trim();
      const description = String(row[10] || "").trim();
      
      if (taskDepartment.toLowerCase() === department.toLowerCase()) {
        tasks.push({
          id: taskId || "TASK_" + (tasks.length + 1),
          name: taskName,
          description: description || taskName,
          assignedTo: assignedTo,
          department: taskDepartment,
          projectName: taskProjectName,
          start: "09:00",
          end: "17:00"
        });
      }
    }
    
    return tasks;
    
  } catch (error) {
    Logger.log("Error in getTasksForDepartment: " + error.toString());
    return [];
  }
}

/**
 * Get all engineers for dropdown selection
 */
function getAllEngineersForDropdown() {
  try {
    const engineers = getAllEngineersWithDetails();
    
    // Return simplified structure for dropdown
    return engineers.map(engineer => ({
      id: engineer.id,
      name: engineer.name,
      department: engineer.department,
      email: engineer.email,
      role: engineer.role,
      level: engineer.level
    }));
    
  } catch (error) {
    Logger.log("Error in getAllEngineersForDropdown: " + error.toString());
    return [];
  }
}

function getAllEngineersWithDetails() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();
    const teamSheet = ss.getSheetByName(config.SHEET_NAMES.TEAM_MAPPING);
    
    if (!teamSheet) {
      return config.DEFAULT_ENGINEERS;
    }
    
    const data = teamSheet.getDataRange().getValues().slice(1);
    const engineers = [];
    
    data.forEach(row => {
      if (!row || row.length === 0) return;
      
      const senior = row[0]?.toString().trim();
      const junior = row[1]?.toString().trim();
      const department = row[2]?.toString().trim();
      const email = row[3]?.toString().trim();
      const role = row[4]?.toString().trim();
      const level = row[5]?.toString().trim();
      
      // Add senior engineers (when they don't have juniors)
      if (senior && senior !== "" && (!junior || junior === "")) {
        engineers.push({
          id: senior.replace(/\s+/g, '').toLowerCase(),
          name: senior,
          role: role || 'Senior',
          department: department || 'Unknown',
          email: email || '',
          level: level || 'Senior'
        });
      }
      
      // Add junior engineers
      if (junior && junior !== "") {
        engineers.push({
          id: junior.replace(/\s+/g, '').toLowerCase(),
          name: junior,
          role: role || 'Engineer',
          department: department || 'Unknown',
          email: email || '',
          level: level || 'Junior'
        });
      }
    });
    
    return engineers.length > 0 ? engineers : config.DEFAULT_ENGINEERS;
  } catch (error) {
    Logger.log("Error in getAllEngineersWithDetails: " + error.toString());
    return getConfig().DEFAULT_ENGINEERS;
  }
}

/**
 * Test engineer data loading
 */
function testEngineerDataLoad() {
  try {
    const engineers = getAllEngineersWithDetails();
    const dropdownEngineers = getAllEngineersForDropdown();
    
    Logger.log("Total engineers: " + engineers.length);
    Logger.log("Dropdown engineers: " + dropdownEngineers.length);
    
    engineers.forEach(eng => {
      Logger.log("Engineer: " + eng.name + " - " + eng.department);
    });
    
    return {
      success: true,
      totalEngineers: engineers.length,
      dropdownEngineers: dropdownEngineers.length,
      sample: engineers.slice(0, 3)
    };
  } catch (error) {
    Logger.log("Error in testEngineerDataLoad: " + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Get engineer by name
 */
function getEngineerByName(engineerName) {
  try {
    const engineers = getAllEngineersWithDetails();
    return engineers.find(eng => 
      eng.name.toLowerCase() === engineerName.toLowerCase() ||
      eng.name.toLowerCase().includes(engineerName.toLowerCase()) ||
      engineerName.toLowerCase().includes(eng.name.toLowerCase())
    );
  } catch (error) {
    Logger.log("Error in getEngineerByName: " + error.toString());
    return null;
  }
}
/**
 * Debug RC Car data for dashboard testing
 */
function debugRCCarData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();
    
    // Get data from all response sheets
    const sheets = [
      { name: config.SHEET_NAMES.EMBEDDED_RESPONSES, dept: "Embedded" },
      { name: config.SHEET_NAMES.AI_RESPONSES, dept: "AI" },
      { name: config.SHEET_NAMES.MECHANICAL_RESPONSES, dept: "Mechanical" }
    ];
    
    const allData = [];
    
    sheets.forEach(cfg => {
      const sh = ss.getSheetByName(cfg.name);
      if (!sh) return;
      
      const data = sh.getDataRange().getValues();
      if (data.length < 2) return;
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row || row.join("").trim() === "") continue;
        
        const ts = row[0] instanceof Date ? row[0] : new Date(row[0]);
        const dateStr = Utilities.formatDate(ts, Session.getScriptTimeZone(), "yyyy-MM-dd");
        const engineer = row[2] || "";
        const project = row[4] || "";
        const taskName = row[1] || "";
        const start = row[7] || "";
        const end = row[8] || "";
        const status = row[9] || "Done";
        
        // Filter for RC Car project specifically
        if (project.toString().toLowerCase().includes("rc car") || 
            project.toString().toLowerCase().includes("rc")) {
          allData.push({
            date: dateStr,
            department: cfg.dept,
            engineer: engineer,
            project: project,
            taskName: taskName,
            start: formatTimeOrDate(start),
            end: formatTimeOrDate(end),
            status: status
          });
        }
      }
    });


    


    // Calculate statistics
    const totalTasks = allData.length;
    const tasksByStatus = {};
    const tasksByDepartment = {};
    
    allData.forEach(item => {
      // Count by status
      tasksByStatus[item.status] = (tasksByStatus[item.status] || 0) + 1;
      
      // Count by department
      tasksByDepartment[item.department] = (tasksByDepartment[item.department] || 0) + 1;
    });
    
    return {
      success: true,
      totalTasks: totalTasks,
      tasksByStatus: tasksByStatus,
      tasksByDepartment: tasksByDepartment,
      rawData: allData.slice(0, 5) // Return first 5 entries as sample
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.toString(),
      totalTasks: 0,
      tasksByStatus: {},
      tasksByDepartment: {},
      rawData: []
    };
  }
}
/**
 * Get projects with deadlines and tasks - UPDATED with header mapping
 */
function getProjectsWithDeadlines() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();
    const sheet = ss.getSheetByName(config.SHEET_NAMES.PROJECTS_TASKS);
    
    if (!sheet) {
      Logger.log("Projects_Tasks sheet not found");
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log("No data in Projects_Tasks sheet");
      return [];
    }
    
    const headers = data[0];
    const projects = {};
    
    // Map headers to indices
    const typeIndex = headers.indexOf('Type');
    const projectIdIndex = headers.indexOf('Project ID');
    const projectNameIndex = headers.indexOf('Project Name');
    const departmentsIndex = headers.indexOf('Department(s)');
    const statusIndex = headers.indexOf('Status');
    const managerIndex = headers.indexOf('Manager');
    const taskIdIndex = headers.indexOf('Task ID');
    const taskNameIndex = headers.indexOf('Task Name');
    const assignedToIndex = headers.indexOf('Assigned To');
    const departmentIndex = headers.indexOf('Department');
    const deadlineIndex = headers.findIndex(header => 
      header.includes('Deadline') || header.includes('Due Date')
    );
    const notesIndex = headers.indexOf('Notes');
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length === 0) continue;
      
      const type = String(row[typeIndex] || "").trim();
      const projectId = String(row[projectIdIndex] || "").trim();
      
      if (!projectId) continue;
      
      if (type === 'Project') {
        projects[projectId] = {
          id: projectId,
          name: String(row[projectNameIndex] || "").trim(),
          departments: String(row[departmentsIndex] || "").trim(),
          status: String(row[statusIndex] || "Active").trim(),
          manager: String(row[managerIndex] || "").trim(),
          deadline: deadlineIndex >= 0 ? String(row[deadlineIndex] || "").trim() : "",
          tasks: []
        };
      } else if (type === 'Task' && projects[projectId]) {
        const taskDeadline = deadlineIndex >= 0 ? String(row[deadlineIndex] || "").trim() : "";
        
        projects[projectId].tasks.push({
          id: String(row[taskIdIndex] || "").trim(),
          name: String(row[taskNameIndex] || "").trim(),
          assignedTo: String(row[assignedToIndex] || "").trim(),
          department: String(row[departmentIndex] || "").trim(),
          status: String(row[statusIndex] || "Not Started").trim(),
          notes: notesIndex >= 0 ? String(row[notesIndex] || "").trim() : "",
          deadline: taskDeadline
        });
      }
    }
    
    // Convert to array and filter out projects with no tasks
    const projectsArray = Object.values(projects).filter(project => project.tasks.length > 0);
    
    Logger.log("Found " + projectsArray.length + " projects with tasks");
    return projectsArray;
    
  } catch (error) {
    Logger.log("Error in getProjectsWithDeadlines: " + error.toString());
    return [];
  }
}

/**
 * Update task deadline and details
 */
function updateTaskDeadline(projectId, taskId, changes) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();
    const projectsSheet = ss.getSheetByName(config.SHEET_NAMES.PROJECTS_TASKS);
    
    if (!projectsSheet) {
      return { success: false, error: "Projects_Tasks sheet not found" };
    }
    
    const data = projectsSheet.getDataRange().getValues();
    let updated = false;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length === 0) continue;
      
      const rowType = String(row[0] || "").trim();
      const currentProjectId = String(row[1] || "").trim();
      const currentTaskId = String(row[6] || "").trim();
      
      if (rowType === "Task" && currentProjectId === projectId && currentTaskId === taskId) {
        // Update the row based on changes
        if (changes.assignedTo !== undefined) {
          projectsSheet.getRange(i + 1, 9).setValue(changes.assignedTo); // Column I - Assigned To
        }
        if (changes.department !== undefined) {
          projectsSheet.getRange(i + 1, 10).setValue(changes.department); // Column J - Department
        }
        if (changes.status !== undefined) {
          // You might want to add a status column or use description
          projectsSheet.getRange(i + 1, 12).setValue("Status: " + changes.status); // Column L - Description
        }
        if (changes.deadline !== undefined) {
          projectsSheet.getRange(i + 1, 11).setValue(changes.deadline); // Column K - Deadline
        }
        
        updated = true;
        break;
      }
    }
    
    if (updated) {
      return { success: true, message: "Task updated successfully" };
    } else {
      return { success: false, error: "Task not found" };
    }
    
  } catch (error) {
    Logger.log("Error in updateTaskDeadline: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Update multiple task deadlines
 */
function updateMultipleTaskDeadlines(changes) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();
    const projectsSheet = ss.getSheetByName(config.SHEET_NAMES.PROJECTS_TASKS);
    
    if (!projectsSheet) {
      return { success: false, error: "Projects_Tasks sheet not found" };
    }
    
    const data = projectsSheet.getDataRange().getValues();
    let updatedCount = 0;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length === 0) continue;
      
      const rowType = String(row[0] || "").trim();
      const currentProjectId = String(row[1] || "").trim();
      const currentTaskId = String(row[6] || "").trim();
      
      if (rowType === "Task" && changes[currentProjectId] && changes[currentProjectId][currentTaskId]) {
        const taskChanges = changes[currentProjectId][currentTaskId];
        
        // Update the row based on changes
        if (taskChanges.assignedTo !== undefined) {
          projectsSheet.getRange(i + 1, 9).setValue(taskChanges.assignedTo);
        }
        if (taskChanges.department !== undefined) {
          projectsSheet.getRange(i + 1, 10).setValue(taskChanges.department);
        }
        if (taskChanges.status !== undefined) {
          projectsSheet.getRange(i + 1, 12).setValue("Status: " + taskChanges.status);
        }
        if (taskChanges.deadline !== undefined) {
          projectsSheet.getRange(i + 1, 11).setValue(taskChanges.deadline);
        }
        
        updatedCount++;
      }
    }
    
    return { success: true, updatedCount: updatedCount };
    
  } catch (error) {
    Logger.log("Error in updateMultipleTaskDeadlines: " + error.toString());
    return { success: false, error: error.toString() };
  }
}
/**
 * Helper function to format time or date
 */
function formatTimeOrDate(value) {
  if (!value) return "";
  
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm");
  }
  
  if (typeof value === 'string' && value.includes(':')) {
    return value; // Already in time format
  }
  
  try {
    const date = new Date(value);
    if (!isNaN(date.getTime())) {
      return Utilities.formatDate(date, Session.getScriptTimeZone(), "HH:mm");
    }
  } catch (e) {
    // Ignore parsing errors
  }
  
  return value.toString();
}

/**
 * Get dashboard data - FIXED for frontend compatibility
 */
function getDashboardData(projectFilter = 'all', statusFilter = 'all') {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();
    
    // Get data from all response sheets
    const sheets = [
      { name: config.SHEET_NAMES.EMBEDDED_RESPONSES, dept: "Embedded" },
      { name: config.SHEET_NAMES.AI_RESPONSES, dept: "AI" },
      { name: config.SHEET_NAMES.MECHANICAL_RESPONSES, dept: "Mechanical" }
    ];
    
    const allData = [];
    const last7Days = getLast7Days();
    
    // Collect data from all sheets
    sheets.forEach(cfg => {
      const sh = ss.getSheetByName(cfg.name);
      if (!sh) return;
      
      const data = sh.getDataRange().getValues();
      if (data.length < 2) return;
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row || row.join("").trim() === "") continue;
        
        try {
          const ts = row[0] instanceof Date ? row[0] : new Date(row[0]);
          const dateStr = Utilities.formatDate(ts, Session.getScriptTimeZone(), "yyyy-MM-dd");
          const dayName = Utilities.formatDate(ts, Session.getScriptTimeZone(), "EEE");
          const shortDate = Utilities.formatDate(ts, Session.getScriptTimeZone(), "MM/dd");
          
          const engineer = row[2] || "";
          const project = row[4] || "";
          const taskName = row[1] || "";
          const start = row[7] || "";
          const end = row[8] || "";
          const status = row[9] || "Done";
          
          // Apply filters
          if (projectFilter !== 'all' && project !== projectFilter) continue;
          if (statusFilter !== 'all' && status !== statusFilter) continue;
          
          // Only include data from last 7 days
          if (last7Days.includes(dateStr)) {
            allData.push({
              date: dateStr,
              day: dayName,
              shortDate: shortDate,
              displayDate: `${dayName}\n${shortDate}`,
              department: cfg.dept,
              engineer: engineer,
              project: project,
              taskName: taskName,
              start: formatTimeOrDate(start),
              end: formatTimeOrDate(end),
              status: status
            });
          }
        } catch (e) {
          continue;
        }
      }
    });
    
    // Calculate statistics
    const totalTasks = allData.length;
    const tasksByStatus = {};
    const tasksByDepartment = {};
    const tasksByDay = {};
    const tasksByEngineer = {};
    
    allData.forEach(item => {
      // Count by status
      tasksByStatus[item.status] = (tasksByStatus[item.status] || 0) + 1;
      
      // Count by department
      tasksByDepartment[item.department] = (tasksByDepartment[item.department] || 0) + 1;
      
      // Count by day with formatted labels
      tasksByDay[item.displayDate] = (tasksByDay[item.displayDate] || 0) + 1;
      
      // Count by engineer
      if (item.engineer) {
        tasksByEngineer[item.engineer] = (tasksByEngineer[item.engineer] || 0) + 1;
      }
    });
    
    // FIXED: Dark theme compatible colors
    const darkThemeColors = [
      '#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7',
      '#D980FA', '#FDA7DF', '#6A89CC', '#82CCDD', '#B8E994'
    ];
    
    // Prepare chart data
    const chartData = {
      totalTasks: totalTasks,
      tasksByStatus: {
        labels: Object.keys(tasksByStatus),
        data: Object.values(tasksByStatus),
        colors: ['#4CAF50', '#FF9800', '#F44336', '#2196F3', '#9C27B0']
      },
      tasksByDepartment: {
        labels: Object.keys(tasksByDepartment),
        data: Object.values(tasksByDepartment),
        colors: ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7']
      },
      tasksByDay: {
        labels: Object.keys(tasksByDay),
        data: Object.values(tasksByDay),
        colors: darkThemeColors.slice(0, Math.min(Object.keys(tasksByDay).length, 7))
      },
      tasksByEngineer: {
        labels: Object.keys(tasksByEngineer),
        data: Object.values(tasksByEngineer),
        // FIXED: Use proper colors for engineer chart
        colors: darkThemeColors.slice(0, Math.min(Object.keys(tasksByEngineer).length, darkThemeColors.length))
      },
      recentActivity: allData.slice(-10).reverse()
    };
    
    return chartData;
    
  } catch (error) {
    Logger.log("Error in getDashboardData: " + error.toString());
    return {
      totalTasks: 0,
      tasksByStatus: { labels: [], data: [], colors: [] },
      tasksByDepartment: { labels: [], data: [], colors: [] },
      tasksByDay: { labels: [], data: [], colors: [] },
      tasksByEngineer: { labels: [], data: [], colors: [] },
      recentActivity: []
    };
  }
}

/**
 * Order chart data by date (not alphabetically)
 */
function orderChartDataByDate(tasksByDay) {
  const dayOrder = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  
  // Convert to array and sort
  const entries = Object.entries(tasksByDay);
  
  entries.sort((a, b) => {
    const dayA = a[0].split('\n')[0];
    const dayB = b[0].split('\n')[0];
    return dayOrder.indexOf(dayA) - dayOrder.indexOf(dayB);
  });
  
  return {
    labels: entries.map(entry => entry[0]),
    data: entries.map(entry => entry[1])
  };
}

/**
 * Fallback chart data in case of errors
 */
function getFallbackChartData() {
  return {
    totalTasks: 0,
    tasksByStatus: { labels: [], data: [], colors: [] },
    tasksByDepartment: { labels: [], data: [], colors: [] },
    tasksByDay: { labels: [], data: [], colors: [] },
    tasksByEngineer: { labels: [], data: [], colors: [] },
    recentActivity: []
  };
}

/**
 * DEBUG: Check what function is being called by Manage Deadline button
 */
function debugManageDeadlineButton() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();
    const sheet = ss.getSheetByName(config.SHEET_NAMES.PROJECTS_TASKS);
    
    if (!sheet) {
      return { error: "Projects_Tasks sheet not found" };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    return {
      success: true,
      totalRows: data.length,
      headers: headers,
      sampleData: data.slice(1, 3), // First 2 data rows
      functionCalled: "getProjectsForDeadlineManagement" // This is what should be called
    };
  } catch (error) {
    return { error: error.toString() };
  }
}
/**
 * Test what the Manage Deadline button should return
 */
function testManageDeadlineButton() {
  const result = getProjectsForDeadlineManagement();
  
  console.log("Manage Deadline Button Test Results:");
  console.log("Projects found: " + result.length);
  
  result.forEach((project, index) => {
    console.log(`${index + 1}. ${project.name} (${project.id}) - ${project.tasks.length} tasks`);
    project.tasks.forEach((task, taskIndex) => {
      console.log(`   - ${task.name} (${task.id}) - Assigned to: ${task.assignedTo}`);
    });
  });
  
  return result;
}

/**
 * Get last 7 days for filtering
 */
function getLast7Days() {
  const days = [];
  for (let i = 6; i >= 0; i--) {
    const date = new Date();
    date.setDate(date.getDate() - i);
    days.push(Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd"));
  }
  return days;
}
