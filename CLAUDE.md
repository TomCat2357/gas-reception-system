# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Google Apps Script project that creates an AppSheet-style data management web application. The app provides CRUD operations on Google Sheets data with a modern web interface.

**Key Architecture:**
- **Code.js**: Main server-side logic handling data operations, web app routing, and spreadsheet management
- **webapp.html**: Modern responsive web application with Material Design-like UI
- **form.html**: In-spreadsheet editing form for direct sheet integration
- **appsscript.json**: Project configuration with webapp and execution API permissions

## Common Development Commands

### Deployment and Publishing
```bash
# Deploy the application using the provided script
./deploy.sh

# Manual deployment commands
clasp push                    # Push code to Google Apps Script
clasp deploy --description "Description"   # Deploy as web app
```

### Testing and Development
- **Direct Testing**: Use `runDirectTest()` function in Code.js to test data operations from Apps Script editor
- **Debug Mode**: Access `?page=debug` parameter in web app URL for debug interface
- **Simple Test**: Use menu "🔧 テスト・デバッグ" → "🔍 簡単テスト画面" in spreadsheet

## Architecture Details

### Data Flow
1. **Web App Route**: `doGet()` handles incoming requests and serves appropriate HTML templates
2. **Data Operations**: Server-side functions (`addNewData`, `getAllData`, `updateData`, `deleteData`) interact with Google Sheets
3. **Sheet Management**: `getDataSheet()` creates or retrieves data sheet with proper headers
4. **Dual Interface**: Both standalone web app and in-spreadsheet form editing

### Key Functions
- `getDataSheet()`: Core function that manages sheet creation/retrieval with error handling
- `createHeaders()`: Sets up standardized column structure with styling
- Data CRUD operations with comprehensive error handling and logging
- Menu system integration for spreadsheet-based operations

### Error Handling Strategy
- Extensive console logging for debugging
- Try-catch blocks with user-friendly error messages
- Timeout handling for long operations
- Graceful fallbacks for missing sheets/data

### UI/UX Patterns
- Mobile-first responsive design
- Loading states and spinners
- Modal dialogs for data entry
- Real-time search/filtering
- Confirmation dialogs for destructive operations

## Data Structure

The application has migrated to reception data management only. Legacy data fields have been removed.

## Development Notes

- **Sheet Management**: Always use `getDataSheet()` to ensure proper sheet initialization
- **Row Indexing**: Remember +2 offset for data operations (header row + 0-based to 1-based conversion)
- **Permission Model**: Configured for "ANYONE" access with "USER_DEPLOYING" execution
- **Timezone**: Set to "Asia/Tokyo" in appsscript.json
- **Runtime**: Uses V8 runtime for modern JavaScript features