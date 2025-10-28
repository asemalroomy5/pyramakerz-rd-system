// ┌───────────────────────────────────────────────┐
// │        PYRAMAKERZ R&D REPORTING SYSTEM        │
// │               Dashboard Service               │
// └───────────────────────────────────────────────┘

/**
 * Get dashboard data with filters - FIXED to handle empty data
 */
function getDashboardData(projectId = "all", statusFilter = "all") {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();
    
    // If no spreadsheet, return empty array
    if (!ss) {
      Logger.log("No active spreadsheet found");
      return [];
    }
    
    const sheets = [
      { name: config.SHEET_NAMES.EMBEDDED_RESPONSES, dept: "Embedded" },
      { name: config.SHEET_NAMES.AI_RESPONSES, dept: "AI" },
      { name: config.SHEET_NAMES.MECHANICAL_RESPONSES, dept: "Mechanical" }
    ];
    
    const allRows = [];
    
    sheets.forEach(cfg => {
      const sh = ss.getSheetByName(cfg.name);
      if (!sh) {
        Logger.log("Sheet not found: " + cfg.name);
        return;
      }
      
      try {
        const data = sh.getDataRange().getValues();
        if (data.length < 2) {
          Logger.log("No data in sheet: " + cfg.name);
          return;
        }
        
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          if (!row || row.join("").trim() === "") continue;
          
          // Your exact column structure from debug
          const ts = row[0] instanceof Date ? row[0] : new Date(row[0]);
          const dateStr = Utilities.formatDate(ts, Session.getScriptTimeZone(), "yyyy-MM-dd");
          const engineer = row[2] || "";           // Column C - Engineer
          const project = row[4] || "";            // Column E - Project
          const taskName = row[1] || "";           // Column B - Task Name
          const start = row[7] || "";              // Column H - Start Time  
          const end = row[8] || "";                // Column I - End Time
          const status = row[9] || "Done";         // Column J - Status
          
          // Apply filters
          const projectMatch = (projectId === "all" || project.toString().toLowerCase().includes(projectId.toLowerCase()));
          const statusMatch = (statusFilter === "all" || status.toString().toLowerCase().includes(statusFilter.toLowerCase()));
          
          if (projectMatch && statusMatch) {
            allRows.push({
              date: dateStr,
              department: cfg.dept,
              engineer: engineer.toString(),
              project: project.toString(),
              taskName: taskName.toString(),
              start: formatTimeOrDate(start),
              end: formatTimeOrDate(end),
              status: status.toString()
            });
          }
        }
      } catch (sheetError) {
        Logger.log("Error processing sheet " + cfg.name + ": " + sheetError);
      }
    });
    
    Logger.log("Retrieved " + allRows.length + " rows of dashboard data");
    return allRows;
  } catch (error) {
    Logger.log("Error in getDashboardData: " + error.toString());
    return [];
  }
}

/**
 * Get dashboard summary statistics - FIXED to always return valid object
 */
function getDashboardSummary() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();
    
    // Default response
    const defaultResponse = { 
      totalSubmissions: 0, 
      uniqueEngineers: 0, 
      activeProjects: 0 
    };
    
    if (!ss) {
      Logger.log("No active spreadsheet");
      return defaultResponse;
    }
    
    const sheets = [
      { name: config.SHEET_NAMES.EMBEDDED_RESPONSES, department: "Embedded" },
      { name: config.SHEET_NAMES.AI_RESPONSES, department: "AI" },
      { name: config.SHEET_NAMES.MECHANICAL_RESPONSES, department: "Mechanical" }
    ];
    
    let total = 0;
    const engineers = new Set();
    
    sheets.forEach(s => {
      const sh = ss.getSheetByName(s.name);
      if (!sh) {
        Logger.log("Sheet not found: " + s.name);
        return;
      }
      
      try {
        const data = sh.getDataRange().getValues();
        if (data.length > 1) {
          const filteredData = data.slice(1).filter(r => r && r.join("").trim() !== "");
          total += filteredData.length;
          filteredData.forEach(row => {
            const engineer = String(row[2] || ""); // Column C - Engineer
            if (engineer && engineer.trim() !== "") {
              engineers.add(engineer.trim());
            }
          });
        }
      } catch (sheetError) {
        Logger.log("Error processing sheet " + s.name + ": " + sheetError);
      }
    });
    
    // Count active projects from Projects_Tasks sheet
    let activeProjects = 0;
    try {
      const projSh = ss.getSheetByName(config.SHEET_NAMES.PROJECTS_TASKS);
      if (projSh) {
        const projData = projSh.getDataRange().getValues();
        for (let i = 1; i < projData.length; i++) {
          const row = projData[i];
          if (!row || row.length === 0) continue;
          
          const rowType = String(row[0] || "").trim();
          const status = String(row[4] || "").trim().toLowerCase(); // Column E - Status
          
          if (rowType === "Project" && (status.includes("active") || status === "")) {
            activeProjects++;
          }
        }
      }
    } catch (projError) {
      Logger.log("Error counting projects: " + projError);
    }
    
    const result = { 
      totalSubmissions: total, 
      uniqueEngineers: engineers.size, 
      activeProjects: activeProjects 
    };
    
    Logger.log("Dashboard summary: " + JSON.stringify(result));
    return result;
  } catch (error) {
    Logger.log("Error in getDashboardSummary: " + error.toString());
    return { 
      totalSubmissions: 0, 
      uniqueEngineers: 0, 
      activeProjects: 0 
    };
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

// ===== NEW DEADLINE MANAGER FUNCTIONS =====

/**
 * Get projects with deadlines and tasks from Project_Tasks sheet
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
    const notesIndex = headers.indexOf('Notes');
    const deadlineIndex = headers.indexOf('Deadline');
    
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
          deadline: String(row[deadlineIndex] || "").trim(),
          tasks: []
        };
      } else if (type === 'Task' && projects[projectId]) {
        projects[projectId].tasks.push({
          id: String(row[taskIdIndex] || "").trim(),
          name: String(row[taskNameIndex] || "").trim(),
          assignedTo: String(row[assignedToIndex] || "").trim(),
          department: String(row[departmentIndex] || "").trim(),
          status: String(row[statusIndex] || "Not Started").trim(),
          notes: String(row[notesIndex] || "").trim(),
          deadline: String(row[deadlineIndex] || "").trim()
        });
      }
    }
    
    // Convert to array and filter projects with tasks
    const projectsArray = Object.values(projects).filter(project => project.tasks.length > 0);
    
    Logger.log("Found " + projectsArray.length + " projects with deadlines");
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
    const sheet = ss.getSheetByName(config.SHEET_NAMES.PROJECTS_TASKS);
    
    if (!sheet) {
      return { success: false, error: "Projects_Tasks sheet not found" };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Map headers to indices
    const typeIndex = headers.indexOf('Type');
    const projectIdIndex = headers.indexOf('Project ID');
    const taskIdIndex = headers.indexOf('Task ID');
    const assignedToIndex = headers.indexOf('Assigned To');
    const departmentIndex = headers.indexOf('Department');
    const statusIndex = headers.indexOf('Status');
    const notesIndex = headers.indexOf('Notes');
    const deadlineIndex = headers.indexOf('Deadline');
    
    // Find and update the task row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length === 0) continue;
      
      const type = String(row[typeIndex] || "").trim();
      const currentProjectId = String(row[projectIdIndex] || "").trim();
      const currentTaskId = String(row[taskIdIndex] || "").trim();
      
      if (type === 'Task' && currentProjectId === projectId && currentTaskId === taskId) {
        // Update the fields
        if (changes.assignedTo !== undefined) {
          sheet.getRange(i + 1, assignedToIndex + 1).setValue(changes.assignedTo);
        }
        if (changes.department !== undefined) {
          sheet.getRange(i + 1, departmentIndex + 1).setValue(changes.department);
        }
        if (changes.status !== undefined) {
          sheet.getRange(i + 1, statusIndex + 1).setValue(changes.status);
        }
        if (changes.notes !== undefined) {
          sheet.getRange(i + 1, notesIndex + 1).setValue(changes.notes);
        }
        if (changes.deadline !== undefined) {
          sheet.getRange(i + 1, deadlineIndex + 1).setValue(changes.deadline);
        }
        
        return { success: true, message: "Task updated successfully" };
      }
    }
    
    return { success: false, error: "Task not found" };
    
  } catch (error) {
    Logger.log("Error in updateTaskDeadline: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Update multiple task deadlines
 */
function updateMultipleTaskDeadlines(edits) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();
    const sheet = ss.getSheetByName(config.SHEET_NAMES.PROJECTS_TASKS);
    
    if (!sheet) {
      return { success: false, error: "Projects_Tasks sheet not found" };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Map headers to indices
    const typeIndex = headers.indexOf('Type');
    const projectIdIndex = headers.indexOf('Project ID');
    const taskIdIndex = headers.indexOf('Task ID');
    const assignedToIndex = headers.indexOf('Assigned To');
    const departmentIndex = headers.indexOf('Department');
    const statusIndex = headers.indexOf('Status');
    const notesIndex = headers.indexOf('Notes');
    const deadlineIndex = headers.indexOf('Deadline');
    
    let updatedCount = 0;
    
    // Process all edits
    for (const projectId in edits) {
      for (const taskId in edits[projectId]) {
        const changes = edits[projectId][taskId];
        
        // Find the task row
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          if (!row || row.length === 0) continue;
          
          const type = String(row[typeIndex] || "").trim();
          const currentProjectId = String(row[projectIdIndex] || "").trim();
          const currentTaskId = String(row[taskIdIndex] || "").trim();
          
          if (type === 'Task' && currentProjectId === projectId && currentTaskId === taskId) {
            // Update the fields
            if (changes.assignedTo !== undefined) {
              sheet.getRange(i + 1, assignedToIndex + 1).setValue(changes.assignedTo);
            }
            if (changes.department !== undefined) {
              sheet.getRange(i + 1, departmentIndex + 1).setValue(changes.department);
            }
            if (changes.status !== undefined) {
              sheet.getRange(i + 1, statusIndex + 1).setValue(changes.status);
            }
            if (changes.notes !== undefined) {
              sheet.getRange(i + 1, notesIndex + 1).setValue(changes.notes);
            }
            if (changes.deadline !== undefined) {
              sheet.getRange(i + 1, deadlineIndex + 1).setValue(changes.deadline);
            }
            
            updatedCount++;
            break;
          }
        }
      }
    }
    
    return { success: true, updatedCount: updatedCount };
    
  } catch (error) {
    Logger.log("Error in updateMultipleTaskDeadlines: " + error.toString());
    return { success: false, error: error.toString() };
  }
}
