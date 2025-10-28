// ┌───────────────────────────────────────────────┐
// │        PYRAMAKERZ R&D REPORTING SYSTEM        │
// │               Project Service                 │
// └───────────────────────────────────────────────┘

/**
 * Submit multiple tasks - UPDATED for engineer system
 */
function submitMultiTasks(payload) {
  if (!payload || !payload.department || !payload.entries) {
    throw new Error("Invalid payload");
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig();
  const dept = String(payload.department).toLowerCase();
  
  let sheetName;
  if (dept.includes("embed")) sheetName = config.SHEET_NAMES.EMBEDDED_RESPONSES;
  else if (dept.includes("ai")) sheetName = config.SHEET_NAMES.AI_RESPONSES;
  else if (dept.includes("mech")) sheetName = config.SHEET_NAMES.MECHANICAL_RESPONSES;
  else if (dept.includes("visual") || dept.includes("graphic")) {
    sheetName = "Visual_Graphic_Responses";
  }
  else sheetName = payload.department.replace(/[\/\\?*[\]]/g, '_') + "_Responses";
  
  const sh = ss.getSheetByName(sheetName);
  if (!sh) {
    const alternativeNames = [
      "Visual_Graphic_Responses",
      "Visual Graphic Responses", 
      "VisualGraphicResponses",
      "Visual-Graphic-Responses"
    ];
    
    for (const altName of alternativeNames) {
      const altSheet = ss.getSheetByName(altName);
      if (altSheet) {
        sheetName = altName;
        break;
      }
    }
    
    if (!ss.getSheetByName(sheetName)) {
      throw new Error("Response sheet not found: " + sheetName);
    }
  }
  
  // UPDATED: Use engineer info from payload if available
  const who = payload.engineerName || (Session.getActiveUser ? Session.getActiveUser().getEmail() : "unknown");
  const ts = payload.selectedDate ? new Date(payload.selectedDate) : new Date();
  const appended = [];
  
  payload.entries.forEach(entry => {
    // Apply auto-time adjustment if times not provided
    const startTime = entry.start || "09:00";
    const endTime = entry.end || "17:00";
    
    const row = [
      ts,
      entry.taskName || "",
      who,
      payload.department,
      entry.projectId || "",
      entry.taskId || "",
      entry.description || "",
      startTime,
      endTime,
      entry.status || "Done"
    ];
    sh.appendRow(row);
    appended.push({ sheet: sheetName, row: row });
  });
  
  // Send notification email
  EmailService.sendTaskSubmissionEmail(payload.department, who, ts, appended);
  
  return { 
    success: true, 
    sheet: sheetName, 
    appended: appended.length,
    engineer: who,
    date: ts
  };
}

/**
 * Submit task from engineer panel with all new features
 */
function submitTaskFromEngineer(engineerData, taskData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();
    
    // Determine if this is a custom task or assigned task
    if (taskData.taskType === 'custom' || !taskData.projectId) {
      return submitCustomTaskForEngineer({
        engineerId: engineerData.id,
        engineerName: engineerData.name,
        projectName: taskData.projectName || "Custom Project",
        taskName: taskData.taskName,
        description: taskData.description || "",
        startTime: taskData.startTime || "09:00",
        endTime: taskData.endTime || "17:00",
        status: taskData.status || "In Progress",
        checklistItems: taskData.checklistItems || ""
      });
    } else {
      // Submit to department response sheet
      const payload = {
        department: engineerData.department,
        engineerName: engineerData.name,
        selectedDate: taskData.selectedDate, // For forgotten date submissions
        entries: [{
          taskName: taskData.taskName,
          projectId: taskData.projectId,
          taskId: taskData.taskId,
          description: taskData.description,
          start: taskData.startTime || "09:00", // Auto-time adjustment
          end: taskData.endTime || "17:00",
          status: taskData.status || "Done"
        }]
      };
      
      return submitMultiTasks(payload);
    }
    
  } catch (error) {
    Logger.log("Error in submitTaskFromEngineer: " + error.toString());
    return { 
      success: false, 
      error: "Task submission failed: " + error.message 
    };
  }
}

/**
 * Update task status with Trello-like system - ENHANCED
 */
function updateTaskStatus(projectId, taskId, newStatus, taskType = 'assigned', engineerId = null) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();
    
    if (taskType === 'custom' && engineerId) {
      // Update custom task status
      const customSheet = ss.getSheetByName(config.SHEET_NAMES.CUSTOM_TASKS);
      if (!customSheet) return { success: false, error: "Custom tasks sheet not found" };
      
      const data = customSheet.getDataRange().getValues();
      const headers = data[0];
      const statusIndex = headers.indexOf('Status');
      const engineerIdIndex = headers.indexOf('Engineer_ID');
      const taskNameIndex = headers.indexOf('Task_Name');
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowEngineerId = String(row[engineerIdIndex] || "").trim();
        const rowTaskName = String(row[taskNameIndex] || "").trim();
        
        if (rowEngineerId === engineerId && rowTaskName === taskId) {
          customSheet.getRange(i + 1, statusIndex + 1).setValue(newStatus);
          return { success: true, message: "Custom task status updated" };
        }
      }
    } else {
      // Update assigned task status in Projects_Tasks sheet
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
          projectsSheet.getRange(i + 1, 5).setValue(newStatus);
          updated = true;
          break;
        }
      }
      
      if (updated) {
        return { success: true, message: "Task status updated" };
      } else {
        return { success: false, error: "Task not found" };
      }
    }
    
    return { success: false, error: "Task not found" };
    
  } catch (error) {
    Logger.log("Error updating task status: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Get tasks with checklist support - NEW
 */
function getTasksWithChecklist(engineerId, projectId = null) {
  try {
    const assignedTasks = getTasksForEngineer(engineerId);
    const customTasks = getCustomTasksForEngineer(engineerId);
    
    // Combine tasks and add checklist structure
    const allTasks = [
      ...assignedTasks.map(task => ({
        ...task,
        type: 'assigned',
        checklist: parseChecklistItems(task.description),
        hasChecklist: task.description && task.description.includes('- [')
      })),
      ...customTasks.map(task => ({
        ...task,
        type: 'custom',
        checklist: parseChecklistItems(task.checklistItems),
        hasChecklist: !!task.checklistItems
      }))
    ];
    
    // Filter by project if specified
    if (projectId) {
      return allTasks.filter(task => task.projectId === projectId || task.projectName === projectId);
    }
    
    return allTasks;
    
  } catch (error) {
    Logger.log("Error in getTasksWithChecklist: " + error.toString());
    return [];
  }
}

/**
 * Parse checklist items from text - NEW
 */
function parseChecklistItems(checklistText) {
  if (!checklistText) return [];
  
  const items = [];
  const lines = checklistText.split('\n');
  
  lines.forEach(line => {
    const trimmed = line.trim();
    if (trimmed.startsWith('- [ ] ') || trimmed.startsWith('- [x] ')) {
      const isChecked = trimmed.startsWith('- [x] ');
      const text = trimmed.substring(6); // Remove "- [ ] " or "- [x] "
      
      items.push({
        text: text,
        checked: isChecked,
        raw: trimmed
      });
    }
  });
  
  return items;
}

/**
 * Update checklist item status - NEW
 */
function updateChecklistItem(engineerId, taskIdentifier, itemIndex, isChecked, taskType = 'custom') {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();
    
    if (taskType === 'custom') {
      const customSheet = ss.getSheetByName(config.SHEET_NAMES.CUSTOM_TASKS);
      if (!customSheet) return { success: false, error: "Custom tasks sheet not found" };
      
      const data = customSheet.getDataRange().getValues();
      const headers = data[0];
      const checklistIndex = headers.indexOf('Checklist_Items');
      const engineerIdIndex = headers.indexOf('Engineer_ID');
      const taskNameIndex = headers.indexOf('Task_Name');
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowEngineerId = String(row[engineerIdIndex] || "").trim();
        const rowTaskName = String(row[taskNameIndex] || "").trim();
        
        if (rowEngineerId === engineerId && rowTaskName === taskIdentifier) {
          const currentChecklist = String(row[checklistIndex] || "");
          const updatedChecklist = updateChecklistInText(currentChecklist, itemIndex, isChecked);
          
          customSheet.getRange(i + 1, checklistIndex + 1).setValue(updatedChecklist);
          return { success: true, message: "Checklist updated" };
        }
      }
    }
    
    return { success: false, error: "Task not found" };
    
  } catch (error) {
    Logger.log("Error updating checklist: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Update checklist in text format - NEW
 */
function updateChecklistInText(checklistText, itemIndex, isChecked) {
  const lines = checklistText.split('\n');
  let currentIndex = 0;
  
  const updatedLines = lines.map(line => {
    const trimmed = line.trim();
    if (trimmed.startsWith('- [ ] ') || trimmed.startsWith('- [x] ')) {
      if (currentIndex === itemIndex) {
        const text = trimmed.substring(6);
        return isChecked ? `- [x] ${text}` : `- [ ] ${text}`;
      }
      currentIndex++;
    }
    return line;
  });
  
  return updatedLines.join('\n');
}

/**
 * Hide completed tasks for engineer - NEW
 */
function hideCompletedTasks(engineerId, hide = true) {
  try {
    // This is mainly a frontend preference, but we can store it
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty(`hideCompleted_${engineerId}`, hide.toString());
    
    return { 
      success: true, 
      message: hide ? "Completed tasks will be hidden" : "All tasks will be shown" 
    };
    
  } catch (error) {
    Logger.log("Error in hideCompletedTasks: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Get hidden tasks preference - NEW
 */
function getHideCompletedPreference(engineerId) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const preference = userProperties.getProperty(`hideCompleted_${engineerId}`);
    return preference === 'true';
  } catch (error) {
    return true; // Default to hide completed
  }
}

/**
 * Submit task with date selection - NEW
 */
function submitTaskWithDateSelection(engineerData, taskData, selectedDate = null) {
  try {
    const submissionDate = selectedDate ? new Date(selectedDate) : new Date();
    
    // Add date to task data
    const taskWithDate = {
      ...taskData,
      selectedDate: submissionDate.toISOString().split('T')[0] // YYYY-MM-DD format
    };
    
    return submitTaskFromEngineer(engineerData, taskWithDate);
    
  } catch (error) {
    Logger.log("Error in submitTaskWithDateSelection: " + error.toString());
    return { 
      success: false, 
      error: "Task submission with date failed: " + error.message 
    };
  }
}

/**
 * Get engineer's recent tasks with date filtering - NEW
 */
function getEngineerRecentTasks(engineerId, daysBack = 7) {
  try {
    const customTasks = getCustomTasksForEngineer(engineerId, false); // Don't hide completed
    
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysBack);
    
    return customTasks.filter(task => {
      const taskDate = task.timestamp instanceof Date ? task.timestamp : new Date(task.timestamp);
      return taskDate >= cutoffDate;
    });
    
  } catch (error) {
    Logger.log("Error in getEngineerRecentTasks: " + error.toString());
    return [];
  }
}

// ┌───────────────────────────────────────────────┐
// │         EXISTING FUNCTIONS (UPDATED)          │
// └───────────────────────────────────────────────┘

/**
 * Generate unique task IDs for a project
 */
function generateTaskIds(projectName, taskCount) {
  const projectAbbr = generateProjectAbbreviation(projectName);
  const taskIds = [];
  
  for (let i = 1; i <= taskCount; i++) {
    taskIds.push(`${projectAbbr}-${String(i).padStart(3, '0')}`);
  }
  
  return taskIds;
}

/**
 * Generate project abbreviation from name
 */
function generateProjectAbbreviation(projectName) {
  const words = projectName.split(' ').filter(word => word.length > 2);
  if (words.length >= 2) {
    return words.slice(0, 2).map(word => word.substring(0, 2)).join('').toUpperCase();
  } else if (words.length === 1) {
    return words[0].substring(0, 3).toUpperCase();
  }
  return 'PROJ';
}

/**
 * Get email recipients for project notifications
 */
function getProjectEmailRecipients(projectData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();
    const teamSheet = ss.getSheetByName(config.SHEET_NAMES.TEAM_MAPPING);
    
    const ctoEmail = projectData.notifications?.ctoEmail || config.CTO_EMAIL;
    const managerEmails = [];
    const assignedEngineerEmails = projectData.engineers?.map(eng => eng.email).filter(email => email) || [];
    
    if (teamSheet && projectData.notifications?.notifyManagers) {
      const teamData = teamSheet.getDataRange().getValues().slice(1);
      
      projectData.departments?.forEach(dept => {
        const seniorEngineers = teamData.filter(row => {
          const department = row[2]?.toString().trim();
          const level = row[5]?.toString().trim();
          const email = row[3]?.toString().trim();
          return department === dept && level === 'Senior' && email;
        });
        
        seniorEngineers.forEach(engineer => {
          const email = engineer[3]?.toString().trim();
          if (email && !managerEmails.includes(email)) {
            managerEmails.push(email);
          }
        });
      });
    }
    
    const allRecipients = [];
    if (projectData.notifications?.notifyCTO && ctoEmail) {
      allRecipients.push(ctoEmail);
    }
    if (projectData.notifications?.notifyManagers) {
      allRecipients.push(...managerEmails);
    }
    if (projectData.notifications?.notifyEngineers) {
      allRecipients.push(...assignedEngineerEmails);
    }
    
    const uniqueRecipients = [...new Set(allRecipients)];
    
    return {
      ctoEmail: ctoEmail,
      managerEmails: managerEmails,
      assignedEngineerEmails: assignedEngineerEmails,
      allRecipients: uniqueRecipients
    };
    
  } catch (error) {
    Logger.log("Error getting email recipients: " + error);
    return {
      ctoEmail: projectData.notifications?.ctoEmail || 'nemr@pyramakerz.com',
      managerEmails: [],
      assignedEngineerEmails: projectData.engineers?.map(eng => eng.email).filter(email => email) || [],
      allRecipients: projectData.engineers?.map(eng => eng.email).filter(email => email) || []
    };
  }
}

/**
 * Create project with approval workflow
 */
function createCTOProjectWithDeadlines(formData) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const config = getConfig();
    const isCTO = (userEmail === config.CTO_EMAIL);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const projectSheet = ss.getSheetByName(config.SHEET_NAMES.PROJECTS_TASKS);
    
    if (!projectSheet) {
      return { success: false, error: "Projects_Tasks sheet not found" };
    }
    
    const projectId = 'PROJ_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'YYYYMMDD_HHmmss');
    
    if (isCTO) {
      return createProjectDirectly(projectId, formData, projectSheet);
    } else {
      return createPendingProject(projectId, formData, userEmail);
    }
    
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

/**
 * Create pending project for approval
 */
function createPendingProject(projectId, formData, creatorEmail) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig();
  
  let pendingSheet = ss.getSheetByName('Pending_Projects');
  if (!pendingSheet) {
    pendingSheet = ss.insertSheet('Pending_Projects');
    pendingSheet.appendRow([
      'Project ID', 'Project Name', 'Description', 'Deadline', 
      'Departments', 'Tasks', 'Created By', 'Created At', 'Status'
    ]);
  }
  
  const taskData = JSON.stringify(formData.tasks);
  const deptData = formData.departments.join(', ');
  
  pendingSheet.appendRow([
    projectId,
    formData.projectName,
    formData.projectDescription,
    formData.projectDeadline,
    deptData,
    taskData,
    creatorEmail,
    new Date(),
    'Pending Review'
  ]);
  
  sendApprovalEmail(projectId, formData, creatorEmail);
  
  return { 
    success: true, 
    projectId: projectId,
    projectName: formData.projectName,
    status: 'pending_approval',
    message: 'Project submitted for CTO approval'
  };
}

/**
 * Send approval email to CTO
 */
function sendApprovalEmail(projectId, formData, creatorEmail) {
  const config = getConfig();
  const subject = `📋 Project Approval Request: ${formData.projectName}`;
  
  const htmlBody = `
    <h2>New Project Approval Request</h2>
    <div style="background: #f8f9fa; padding: 20px; border-radius: 8px;">
      <h3>${formData.projectName}</h3>
      <p><strong>Created by:</strong> ${creatorEmail}</p>
      <p><strong>Description:</strong> ${formData.projectDescription}</p>
      <p><strong>Deadline:</strong> ${formData.projectDeadline}</p>
      <p><strong>Departments:</strong> ${formData.departments.join(', ')}</p>
      <p><strong>Tasks:</strong> ${formData.tasks.length} tasks</p>
    </div>
    
    <div style="margin-top: 20px;">
      <a href="${ScriptApp.getService().getUrl()}?action=approve&id=${projectId}" 
         style="background: #28a745; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; margin-right: 10px;">
        ✅ Approve
      </a>
      <a href="${ScriptApp.getService().getUrl()}?action=edit&id=${projectId}" 
         style="background: #ffc107; color: black; padding: 10px 20px; text-decoration: none; border-radius: 5px;">
        ✏️ Edit & Approve
      </a>
    </div>
  `;
  
  MailApp.sendEmail({
    to: config.CTO_EMAIL,
    subject: subject,
    htmlBody: htmlBody
  });
}

/**
 * Quick Project Creator - Simple version for manager dashboard
 */
function quickCreateProject(projectData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();
    const projectsSheet = ss.getSheetByName(config.SHEET_NAMES.PROJECTS_TASKS);
    
    if (!projectsSheet) {
      throw new Error("Projects_Tasks sheet not found");
    }
    
    const projectId = projectData.projectName.replace(/[^a-zA-Z]/g, '').substring(0, 3).toUpperCase();
    
    const projectRow = [
      'Project',
      projectId,
      projectData.projectName,
      projectData.department,
      'Active',
      'CTO',
      '', '', '', '',
      projectData.description || '',
      projectData.deadline
    ];
    
    projectsSheet.appendRow(projectRow);
    
    let tasksCreated = 0;
    const taskCount = projectData.taskCount || 1;
    
    for (let i = 1; i <= taskCount; i++) {
      const taskId = `${projectId}-${String(i).padStart(3, '0')}`;
      const taskRow = [
        'Task',
        projectId,
        projectData.projectName,
        '', '', '',
        taskId,
        `Task ${i} - ${projectData.projectName}`,
        projectData.assignedTo || '',
        projectData.department,
        `Task ${i} for ${projectData.projectName}`,
        projectData.deadline
      ];
      
      projectsSheet.appendRow(taskRow);
      tasksCreated++;
    }
    
    return {
      success: true,
      projectId: projectId,
      tasksCreated: tasksCreated
    };
    
  } catch (error) {
    Logger.log("Error in quickCreateProject: " + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Get projects for deadline management - FIXED COLUMN MAPPING
 */
function getProjectsForDeadlineManagement() {
  try {
    Logger.log("=== STARTING getProjectsForDeadlineManagement ===");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getConfig();
    const sheet = ss.getSheetByName(config.SHEET_NAMES.PROJECTS_TASKS);
    
    if (!sheet) {
      Logger.log("ERROR: Projects_Tasks sheet not found");
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    Logger.log("Total rows in sheet: " + data.length);
    
    if (data.length < 2) {
      Logger.log("No data in Projects_Tasks sheet");
      return [];
    }
    
    const projects = {};
    let projectCount = 0;
    let taskCount = 0;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length === 0) continue;
      
      const rowType = String(row[0] || "").trim();
      const projectId = String(row[1] || "").trim();
      
      if (!projectId) continue;
      
      if (rowType === 'Project') {
        projectCount++;
        projects[projectId] = {
          id: projectId,
          name: String(row[2] || "").trim(),
          departments: String(row[3] || "").trim(),
          status: String(row[4] || "Active").trim(),
          manager: String(row[5] || "").trim(),
          deadline: row[11] ? String(row[11] || "").trim() : "",
          tasks: []
        };
      } 
      else if (rowType === 'Task' && projects[projectId]) {
        taskCount++;
        projects[projectId].tasks.push({
          id: String(row[6] || "").trim(),
          name: String(row[7] || "").trim(),
          assignedTo: String(row[8] || "").trim(),
          department: String(row[9] || "").trim(),
          deadline: row[11] ? String(row[11] || "").trim() : "",
          status: "Not Started",
          description: String(row[10] || "").trim()
        });
      }
    }
    
    const projectsArray = Object.values(projects);
    
    Logger.log(`=== COMPLETED: Found ${projectsArray.length} projects with ${taskCount} total tasks ===`);
    
    return projectsArray;
    
  } catch (error) {
    Logger.log("ERROR in getProjectsForDeadlineManagement: " + error.toString());
    return [];
  }
}

/**
 * COMPATIBILITY FUNCTION: Make Manage Deadline button work
 */
function getProjectsWithDeadlines() {
  Logger.log("getProjectsWithDeadlines() called - redirecting to getProjectsForDeadlineManagement()");
  return getProjectsForDeadlineManagement();
}

/**
 * Update task deadline and details - FIXED COLUMN MAPPING
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
        if (changes.assignedTo !== undefined) {
          projectsSheet.getRange(i + 1, 9).setValue(changes.assignedTo);
        }
        if (changes.department !== undefined) {
          projectsSheet.getRange(i + 1, 10).setValue(changes.department);
        }
        if (changes.status !== undefined) {
          projectsSheet.getRange(i + 1, 5).setValue(changes.status);
        }
        if (changes.deadline !== undefined) {
          projectsSheet.getRange(i + 1, 12).setValue(changes.deadline);
        }
        if (changes.description !== undefined) {
          projectsSheet.getRange(i + 1, 11).setValue(changes.description);
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
        
        if (taskChanges.assignedTo !== undefined) {
          projectsSheet.getRange(i + 1, 9).setValue(taskChanges.assignedTo);
        }
        if (taskChanges.department !== undefined) {
          projectsSheet.getRange(i + 1, 10).setValue(taskChanges.department);
        }
        if (taskChanges.status !== undefined) {
          projectsSheet.getRange(i + 1, 5).setValue(taskChanges.status);
        }
        if (taskChanges.deadline !== undefined) {
          projectsSheet.getRange(i + 1, 12).setValue(taskChanges.deadline);
        }
        if (taskChanges.description !== undefined) {
          projectsSheet.getRange(i + 1, 11).setValue(taskChanges.description);
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

// ┌───────────────────────────────────────────────┐
// │           CENTRALIZED EMAIL SERVICE           │
// └───────────────────────────────────────────────┘

/**
 * Centralized email service - ALWAYS sends from authorized account
 */
const EmailService = {
  /**
   * Send project creation email (always from authorized account)
   */
  sendProjectCreationEmail(projectData, projectId, recipients) {
    try {
      const subject = `🚀 New Project Created: ${projectData.projectName}`;
      const body = this.createProjectEmailBody(projectData, projectId);
      
      Logger.log(`Sending project email to: ${recipients.join(', ')}`);
      
      MailApp.sendEmail({
        to: recipients.join(','),
        subject: subject,
        htmlBody: body
      });
      
      return { success: true, sentTo: recipients.length };
      
    } catch (error) {
      Logger.log("EmailService error: " + error.toString());
      return { success: false, error: error.toString() };
    }
  },
  
  /**
   * Send task submission notification
   */
  sendTaskSubmissionEmail(department, submitter, timestamp, submissions) {
    try {
      const config = getConfig();
      const emails = [config.SEIF_EMAIL, config.CTO_EMAIL];
      const subj = `New ${department} daily report submissions (${submissions.length})`;
      
      let body = `New submissions by ${submitter} on ${Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm")}\n`;
      
      submissions.forEach((a, i) => {
        body += `${i + 1}. Project: ${a.row[4]} | Task: ${a.row[5]} | Title: ${a.row[1]}\n   Desc: ${a.row[6]}\n   Status: ${a.row[9]}\n`;
      });
      
      MailApp.sendEmail({
        to: emails.join(","), 
        subject: subj, 
        body: body 
      });
      
      return { success: true };
    } catch (e) {
      Logger.log("EmailService submission error: " + e);
      return { success: false, error: e.toString() };
    }
  },
  
  createProjectEmailBody(projectData, projectId) {
    const tasksList = projectData.tasks.map(task => 
      `<li><strong>${task.id}:</strong> ${task.name} (Assigned to: ${task.assignedTo || 'Unassigned'})</li>`
    ).join('');
    
    return `
      <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
        <div style="background: linear-gradient(135deg, #ff6a00, #ff8c42); color: white; padding: 20px; text-align: center; border-radius: 10px 10px 0 0;">
          <h1>🚀 New Project Created</h1>
          <h2>${projectData.projectName}</h2>
          <p>Project ID: ${projectId}</p>
        </div>
        
        <div style="padding: 20px; background: #f8f9fa; border-radius: 0 0 10px 10px;">
          <div style="background: white; padding: 15px; margin-bottom: 15px; border-radius: 8px; border-left: 4px solid #ff6a00;">
            <h3>📋 Project Details</h3>
            <p><strong>Description:</strong> ${projectData.projectDescription}</p>
            <p><strong>Deadline:</strong> ${projectData.projectDeadline}</p>
            <p><strong>Departments:</strong> ${projectData.departments.join(', ')}</p>
            <p><strong>Status:</strong> Active</p>
          </div>
          
          <div style="background: white; padding: 15px; margin-bottom: 15px; border-radius: 8px; border-left: 4px solid #28a745;">
            <h3>📝 Project Tasks</h3>
            <ul>${tasksList}</ul>
          </div>
          
          <div style="background: white; padding: 15px; border-radius: 8px; border-left: 4px solid #007bff;">
            <h3>👥 Team</h3>
            <p><strong>Assigned Engineers:</strong> ${projectData.engineers.map(eng => eng.name).join(', ') || 'None'}</p>
          </div>
        </div>
      </div>
    `;
  }
};

/**
 * Test the engineer integration
 */
function testEngineerIntegration() {
  try {
    // Test authentication
    const authResult = authenticateEngineer("Eng. Seif", "password123");
    
    if (!authResult.success) {
      return { success: false, error: "Authentication failed" };
    }
    
    const engineer = authResult.engineer;
    
    // Test assigned projects
    const projects = getAssignedProjectsForEngineer(engineer.name, engineer.department);
    
    // Test custom tasks
    const customTasks = getCustomTasksForEngineer(engineer.id);
    
    // Test task submission
    const testTask = {
      taskType: 'custom',
      taskName: 'Test Integration Task',
      description: 'Testing the engineer integration',
      startTime: '09:00',
      endTime: '17:00',
      status: 'In Progress'
    };
    
    const submissionResult = submitTaskFromEngineer(engineer, testTask);
    
    return {
      success: true,
      engineer: engineer,
      assignedProjects: projects.length,
      customTasks: customTasks.length,
      taskSubmission: submissionResult.success
    };
    
  } catch (error) {
    return {
      success: false,
      error: "Integration test failed: " + error.toString()
    };
  }
}