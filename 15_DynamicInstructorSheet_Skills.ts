/**
 * YSL Hub v2 Dynamic Instructor Sheet Skills Module
 * 
 * This module contains the functions related to skills management
 * in the dynamic instructor sheet.
 * 
 * @author Sean R. Sullivan
 * @version 1.0
 * @date 2025-05-14
 */

/**
 * Gets all skill headers from the Swimmer Records Workbook
 * @return Object with skill headers categorized by type
 */
function getSkillsFromSwimmerRecords() {
  try {
    // Get Swimmer Records URL from different possible sources with robust error handling
    let swimmerRecordsUrl = null;
    
    // Try different ways to get the URL
    try {
      // First, try GlobalFunctions if available
      if (typeof GlobalFunctions !== 'undefined' && typeof GlobalFunctions.safeGetProperty === 'function') {
        swimmerRecordsUrl = GlobalFunctions.safeGetProperty('swimmerRecordsUrl');
        
        // If not found with direct property name, try through CONFIG object
        if (!swimmerRecordsUrl && typeof CONFIG !== 'undefined' && CONFIG.SWIMMER_RECORDS_URL) {
          swimmerRecordsUrl = GlobalFunctions.safeGetProperty(CONFIG.SWIMMER_RECORDS_URL);
        }
      }
      
      // If still not found, try getting from AdministrativeModule
      if (!swimmerRecordsUrl && typeof AdministrativeModule !== 'undefined' && 
          typeof AdministrativeModule.getSystemConfiguration === 'function') {
        const config = AdministrativeModule.getSystemConfiguration();
        if (config && config.swimmerRecordsUrl) {
          swimmerRecordsUrl = config.swimmerRecordsUrl;
        }
      }
      
      // Last resort - direct property access
      if (!swimmerRecordsUrl) {
        swimmerRecordsUrl = PropertiesService.getScriptProperties().getProperty('swimmerRecordsUrl');
      }
    } catch (propError) {
      Logger.log(`Error accessing configuration: ${propError.message}`);
      // Don't return yet, continue with hardcoded ID as last resort
    }
    
    // If URL not found from any source, try a known hardcoded ID for testing
    if (!swimmerRecordsUrl) {
      Logger.log('Swimmer Records URL not found in system configuration. Using fallback skills.');
      return createFallbackSkills();
    }
    
    // Extract spreadsheet ID from URL with multiple fallback methods
    let ssId = null;
    
    try {
      // Try with GlobalFunctions first
      if (typeof GlobalFunctions !== 'undefined' && typeof GlobalFunctions.extractIdFromUrl === 'function') {
        ssId = GlobalFunctions.extractIdFromUrl(swimmerRecordsUrl);
      }
      
      // If that fails, try direct regex extraction
      if (!ssId) {
        const urlPattern = /[-\w]{25,}/;
        const match = swimmerRecordsUrl.match(urlPattern);
        ssId = match ? match[0] : null;
      }
      
      // If that fails too, just use the URL as-is
      if (!ssId) {
        ssId = swimmerRecordsUrl;
      }
    } catch (urlError) {
      Logger.log(`Error extracting ID from URL: ${urlError.message}`);
      // Try using URL directly
      ssId = swimmerRecordsUrl;
    }
    
    if (!ssId) {
      Logger.log('Invalid Swimmer Records URL - could not extract valid ID');
      return createFallbackSkills();
    }
    
    // Log the spreadsheet ID we're trying to open
    Logger.log(`Attempting to open Swimmer Records with ID: ${ssId}`);
    
    try {
      // Try to open the Swimmer Records Workbook with careful error handling
      let swimmerSS = null;
      
      try {
        // Try using GlobalFunctions for safer access if available
        if (typeof GlobalFunctions !== 'undefined' && typeof GlobalFunctions.safeGetSpreadsheetById === 'function') {
          swimmerSS = GlobalFunctions.safeGetSpreadsheetById(ssId);
        } else {
          // Direct access as fallback
          swimmerSS = SpreadsheetApp.openById(ssId);
        }
      } catch (accessError) {
        const errorMsg = `Error accessing Swimmer Records: ${accessError.message}`;
        Logger.log(errorMsg);
        
        // Log with ErrorHandling if available
        if (typeof ErrorHandling !== 'undefined' && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(errorMsg, 'ERROR', 'getSkillsFromSwimmerRecords');
        }
        
        // Fall back to test data
        return createFallbackSkills();
      }
      
      if (!swimmerSS) {
        Logger.log('Could not open Swimmer Records spreadsheet - null result');
        return createFallbackSkills();
      }
      
      // Find the right sheet - either Skills sheet or first sheet
      let swimmerSheet = null;
      try {
        // Try to get Skills sheet first
        swimmerSheet = swimmerSS.getSheetByName('Skills');
        
        // If not found, try first sheet
        if (!swimmerSheet) {
          const sheets = swimmerSS.getSheets();
          if (sheets && sheets.length > 0) {
            swimmerSheet = sheets[0];
          }
        }
      } catch (sheetError) {
        Logger.log(`Error finding appropriate sheet: ${sheetError.message}`);
        return createFallbackSkills();
      }
      
      if (!swimmerSheet) {
        Logger.log('No suitable sheet found in Swimmer Records Workbook');
        return createFallbackSkills();
      }
      
      // Get the header row with error handling
      let headerRow = [];
      try {
        headerRow = swimmerSheet.getRange(1, 1, 1, swimmerSheet.getLastColumn()).getValues()[0];
      } catch (rangeError) {
        Logger.log(`Error reading header row: ${rangeError.message}`);
        return createFallbackSkills();
      }
      
      if (!headerRow || headerRow.length === 0) {
        Logger.log('Empty header row in Swimmer Records workbook');
        return createFallbackSkills();
      }
      
      // Log the headers for debugging
      Logger.log(`Swimmer Records headers: ${JSON.stringify(headerRow)}`);
      
      // Categorize skills
      const skills = {
        stage: [], // For stage skills (prefixed with S1, S2, etc.)
        saw: []    // For SAW skills (prefixed with SAW)
      };
      
      // Start from column 3 (after first and last name)
      for (let i = 2; i < headerRow.length; i++) {
        const header = headerRow[i];
        if (!header) continue;
        
        // Check skill type by prefix
        const headerStr = header.toString();
        if (headerStr.startsWith('S') && !headerStr.startsWith('SAW')) {
          skills.stage.push({
            index: i,
            header: headerStr,
            description: '' // Add required description field
          });
        } else if (headerStr.startsWith('SAW')) {
          skills.saw.push({
            index: i,
            header: headerStr,
            description: '' // Add required description field
          });
        }
      }
      
      // Log the skills we found
      Logger.log(`Found ${skills.stage.length} stage skills and ${skills.saw.length} SAW skills`);
      
      // If no skills found, fall back to test skills
      if (skills.stage.length === 0 && skills.saw.length === 0) {
        Logger.log('No skills found in Swimmer Records, using fallback skills');
        return createFallbackSkills();
      }
      
      return skills;
      
    } catch (finalError) {
      const errorMsg = `Failed to process Swimmer Records: ${finalError.message}`;
      Logger.log(errorMsg);
      
      // Log with ErrorHandling if available
      if (typeof ErrorHandling !== 'undefined' && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage(errorMsg, 'ERROR', 'getSkillsFromSwimmerRecords');
      }
      
      return createFallbackSkills();
    }
  } catch (error) {
    Logger.log(`Error getting skills from Swimmer Records: ${error.message}`);
    // Return fallback skills instead of throwing an error
    return createFallbackSkills();
  }
}

/**
 * Creates fallback skills for testing when Swimmer Records is unavailable
 * @return Object with test skill headers
 */
function createFallbackSkills() {
  const skills = {
    stage: [],
    saw: []
  };
  
  // Add some stage skills for testing
  const stageNames = ['S1', 'S2', 'S3', 'S4', 'S5', 'S6'];
  const skillTypes = ['Float', 'Kick', 'Submerge', 'Arm Strokes', 'Breathing'];
  
  let index = 2; // Start after first and last name columns
  
  // Add stage skills
  for (const stage of stageNames) {
    for (const skill of skillTypes) {
      skills.stage.push({
        index: index++,
        header: `${stage} ${skill}`,
        stage: stage.replace('S', ''), // Extract numeric stage value
        description: `Skill test for ${stage} ${skill}` // Add required description
      });
    }
  }
  
  // Add SAW skills
  const sawSkills = ['SAW Water Safety', 'SAW Life Jacket', 'SAW Help Others', 'SAW Call for Help'];
  for (const skill of sawSkills) {
    skills.saw.push({
      index: index++,
      header: skill,
      description: `Safety skill: ${skill}` // Add required description
    });
  }
  
  Logger.log(`Created ${skills.stage.length} fallback stage skills and ${skills.saw.length} fallback SAW skills`);
  return skills;
}

/**
 * Extracts stage information from a class name
 * @param className - The class name to analyze
 * @return Stage information object with value and prefix
 */
function extractStageFromClassName(className) {
  if (!className) return { value: '', prefix: '' };
  
  // Normalize the class name for consistent parsing
  const normalizedName = className.toLowerCase().trim();
  
  // Pattern 1: "Stage 1" or "Stage A"
  let stageMatch = normalizedName.match(/stage\s+([1-6a-f])/i);
  if (stageMatch && stageMatch[1]) {
    return { 
      value: stageMatch[1],
      prefix: 'S'
    };
  }
  
  // Pattern 2: "S1" or "SA"
  stageMatch = normalizedName.match(/\bs([1-6a-f])\b/i);
  if (stageMatch && stageMatch[1]) {
    return { 
      value: stageMatch[1],
      prefix: 'S'
    };
  }
  
  // Pattern 3: Look for "X" where X is a digit 1-6 or letter A-F that might be stage
  // Only use this if it's likely to be referring to a stage
  if (normalizedName.includes('swim') || 
      normalizedName.includes('aqua') || 
      normalizedName.includes('water')) {
    stageMatch = normalizedName.match(/\b([1-6a-f])\b/i);
    if (stageMatch && stageMatch[1]) {
      return { 
        value: stageMatch[1],
        prefix: 'S'
      };
    }
  }
  
  // For backward compatibility, return an empty string when no stage is found
  return { value: '', prefix: '' };
}

/**
 * Filters skills by stage based on the class name
 * @param allSkills - The complete skills object 
 * @param stageInfo - Stage info with value and prefix
 * @return Filtered skills object
 */
function filterSkillsByStage(allSkills, stageInfo) {
  // If no stage specified or no skills available, return all skills
  if (!stageInfo || !stageInfo.value || !allSkills) {
    return allSkills;
  }
  
  const stageValue = stageInfo.value.toString().toLowerCase();
  const stagePrefix = stageInfo.prefix || 'S';
  const stageCode = `${stagePrefix}${stageValue}`;
  
  Logger.log(`Filtering skills to show only stage ${stageCode} (no previous stages)`);
  
  const result = {
    stage: [],
    saw: allSkills.saw || [] // Keep all SAW skills
  };
  
  // Only include skills for the specified stage and prior stages
  if (allSkills.stage && allSkills.stage.length > 0) {
    for (const skill of allSkills.stage) {
      // Extract the stage from the skill header (e.g., 'S1 Float' → 'S1')
      const skillStageInfo = extractStageFromSkillHeader(skill.header);
      
      // MODIFIED: Include only current stage skills, not previous stage
      if (skillStageInfo === stageCode) {
        result.stage.push(skill);
      }
    }
  }
  
  Logger.log(`Filtered ${result.stage.length} stage skills and kept ${result.saw.length} SAW skills`);
  
  // If we didn't find any skills for this stage, return all skills
  if (result.stage.length === 0) {
    Logger.log('No skills found for specified stage, returning all skills');
    return allSkills;
  }
  
  return result;
}

/**
 * Extracts stage code from a skill header
 * @param header - The skill header
 * @return The complete stage code (e.g., 'S1', 'SA') or empty string
 */
function extractStageFromSkillHeader(header) {
  if (!header) return '';
  
  // Check for common patterns like 'S1' or 'SA' at the beginning
  const match = header.toString().match(/^(S[1-6A-Fa-f])\s/);
  if (match && match[1]) {
    return match[1].toUpperCase();
  }
  
  return '';
}