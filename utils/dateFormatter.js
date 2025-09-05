/**
 * Date Formatting Utilities
 * Provides consistent date formatting across the application
 */

/**
 * Format date to human-readable format for Excel
 * @param {Date} date - Date object to format (defaults to current date)
 * @returns {string} Formatted date string (e.g., "Sep 5, 2025 10:32 AM")
 */
function formatDateForExcel(date = new Date()) {
    const options = {
        year: 'numeric',
        month: 'short', 
        day: 'numeric',
        hour: 'numeric',
        minute: '2-digit',
        hour12: true,
        timeZone: 'UTC'
    };
    
    return date.toLocaleDateString('en-US', options);
}

/**
 * Format date to short readable format
 * @param {Date} date - Date object to format (defaults to current date)  
 * @returns {string} Short formatted date (e.g., "Sep 5, 2025")
 */
function formatDateShort(date = new Date()) {
    const options = {
        year: 'numeric',
        month: 'short',
        day: 'numeric',
        timeZone: 'UTC'
    };
    
    return date.toLocaleDateString('en-US', options);
}

/**
 * Format date to readable format with timezone
 * @param {Date} date - Date object to format (defaults to current date)
 * @returns {string} Formatted date with timezone (e.g., "Sep 5, 2025 10:32 AM UTC")
 */
function formatDateWithTimezone(date = new Date()) {
    const options = {
        year: 'numeric',
        month: 'short',
        day: 'numeric', 
        hour: 'numeric',
        minute: '2-digit',
        hour12: true,
        timeZone: 'UTC',
        timeZoneName: 'short'
    };
    
    return date.toLocaleDateString('en-US', options);
}

/**
 * Get current formatted date for Excel 'Last Updated' field
 * @returns {string} Current date in readable format
 */
function getCurrentFormattedDate() {
    return formatDateForExcel(new Date());
}

module.exports = {
    formatDateForExcel,
    formatDateShort,
    formatDateWithTimezone,
    getCurrentFormattedDate
};