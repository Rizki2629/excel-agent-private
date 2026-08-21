const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

console.log('========================================');
console.log('Excel Agent - Add-in Installer');
console.log('========================================');
console.log('');

// Check if running on Windows
if (process.platform !== 'win32') {
    console.log('This script only works on Windows.');
    process.exit(1);
}

// Office add-in registry path
const officePath = 'HKCU\\Software\\Microsoft\\Office\\16.0\\Excel\\Options';

console.log('Installing Excel Agent Add-in...');
console.log('');

try {
    // Method 1: Use Manfred's Office Add-in CLI (if installed)
    console.log('Trying to use office-addin CLI...');
    execSync('office-addin sideload install manifest.xml', { 
        stdio: 'inherit',
        cwd: path.join(__dirname, '..')
    });
    console.log('');
    console.log('Add-in installed successfully!');
    
} catch (error) {
    console.log('');
    console.log('office-addin CLI not found. Using manual installation...');
    console.log('');
    console.log('========================================');
    console.log('Manual Installation:');
    console.log('========================================');
    console.log('');
    console.log('1. Start the server:');
    console.log('   npm start');
    console.log('');
    console.log('2. Open Excel');
    console.log('');
    console.log('3. Go to: File > Options > Trust Center > Trust Center Settings');
    console.log('');
    console.log('4. Click "Trusted Add-in Catalogs"');
    console.log('');
    console.log('5. Add this URL:');
    console.log('   https://localhost:3000');
    console.log('');
    console.log('6. Click "Add catalog"');
    console.log('');
    console.log('7. Go to: Insert > Office Add-ins > My Add-ins');
    console.log('');
    console.log('8. Click "Shared Folder" tab');
    console.log('');
    console.log('9. Click "Add" next to Excel Agent');
    console.log('');
    console.log('========================================');
    console.log('');
    console.log('Alternative: Use the Excel Agent.xlsm file');
    console.log('which contains the VBA-based version.');
}
