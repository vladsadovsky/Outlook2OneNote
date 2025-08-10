#!/usr/bin/env node

/**
 * Environment Configuration Validator
 * 
 * This script validates that the .env file is properly configured
 * and prov  // Summary
  console.log(colorize('📊 Validation Summary:', 'cyan'));
  
  if (configIsValid && warnings.length === 0) {
    console.log(colorize('✅ All required environment variables are properly configured!', 'green'));
    console.log(colorize('🚀 Your application should be ready to run.', 'green'));
  } else if (configIsValid && warnings.length > 0) {
    console.log(colorize('⚠️  Configuration is functional but has placeholder values:', 'yellow'));
    warnings.forEach(varName => {
      console.log(colorize(`   - ${varName}`, 'yellow'));
    });
    console.log(colorize('\n💡 Update these with real values for production use.', 'yellow'));
  } else {
    console.log(colorize('❌ Environment configuration has errors that must be fixed:', 'red'));rror messages for missing or invalid values.
 * 
 * Usage:
 *   node scripts/validate-env.js
 */

const fs = require('fs');
const path = require('path');

// ANSI color codes for console output
const colors = {
  reset: '\x1b[0m',
  red: '\x1b[31m',
  green: '\x1b[32m',
  yellow: '\x1b[33m',
  blue: '\x1b[34m',
  magenta: '\x1b[35m',
  cyan: '\x1b[36m'
};

function colorize(text, color) {
  return colors[color] + text + colors.reset;
}

function validateEnvFile() {
  const envPath = path.join(process.cwd(), '.env');
  const templatePath = path.join(process.cwd(), '.env-template');
  
  console.log(colorize('\n🔍 Environment Configuration Validator\n', 'cyan'));
  
  // Check if .env file exists
  if (!fs.existsSync(envPath)) {
    console.log(colorize('❌ .env file not found!', 'red'));
    
    if (fs.existsSync(templatePath)) {
      console.log(colorize('\n💡 Solution:', 'yellow'));
      console.log('   Copy the template file to create your .env:');
      console.log(colorize('   cp .env-template .env', 'blue'));
      console.log('   Then update the values with your actual Azure AD configuration.\n');
    } else {
      console.log(colorize('❌ .env-template file also not found!', 'red'));
      console.log('   Please check if you\'re in the correct project directory.\n');
    }
    
    return false;
  }
  
  // Load environment variables
  require('dotenv').config();
  
  const requiredVars = {
    'CLIENT_ID': {
      required: true,
      example: 'your-azure-app-client-id-here',
      description: 'Azure AD Application (Client) ID',
      validate: (value) => value && value.match(/^[0-9a-f-]{36}$/)
    },
    'CLIENT_SECRET': {
      required: false, // Optional for client-side PKCE
      example: 'your-secret-here',
      description: 'Azure AD Application Client Secret (for backend services)',
      validate: (value) => !value || value.length > 10 // Either empty or substantial length
    },
    'TENANT_ID': {
      required: true,
      example: 'common',
      description: 'Azure AD Tenant ID or "common"',
      validate: (value) => value === 'common' || value.match(/^[0-9a-f-]{36}$/)
    },
    'REDIRECT_URI': {
      required: true,
      example: 'https://localhost:3000/auth/callback',
      description: 'OAuth redirect URI',
      validate: (value) => value && value.startsWith('https://')
    },
    'AUTHORITY': {
      required: true,
      example: 'https://login.microsoftonline.com/common',
      description: 'Azure AD authority URL',
      validate: (value) => value && value.startsWith('https://login.microsoftonline.com/')
    },
    'GRAPH_SCOPES': {
      required: true,
      example: 'https://graph.microsoft.com/Notes.Read https://graph.microsoft.com/User.Read',
      description: 'Microsoft Graph API scopes',
      validate: (value) => value && value.includes('graph.microsoft.com')
    }
  };
  
  let configIsValid = true;
  let warnings = [];
  
  console.log(colorize('📋 Validating required environment variables:\n', 'blue'));
  
  for (const [varName, config] of Object.entries(requiredVars)) {
    const value = process.env[varName];
    const hasValue = value && value.trim() !== '';
    const isPlaceholder = value && (
      value.includes('your-') || 
      value.includes('here') || 
      value.includes('placeholder') ||
      (config.example && value === config.example && 
       !['TENANT_ID', 'AUTHORITY', 'GRAPH_SCOPES'].includes(varName)) // These can match examples legitimately
    );
    const isValid = hasValue && config.validate ? config.validate(value) : hasValue;
    
    if (config.required && (!hasValue || isPlaceholder || !isValid)) {
      console.log(colorize(`❌ ${varName}`, 'red'));
      console.log(`   ${config.description}`);
      console.log(colorize(`   Expected format: ${config.example}`, 'yellow'));
      
      if (isPlaceholder) {
        console.log(colorize('   ⚠️  Still using placeholder value - please update with real value', 'yellow'));
      } else if (!hasValue) {
        console.log(colorize('   ❌ Missing or empty', 'red'));
      } else if (!isValid) {
        console.log(colorize('   ❌ Invalid format', 'red'));
        console.log(colorize(`   Current value: ${value}`, 'yellow'));
      }
      
      configIsValid = false;
    } else if (hasValue && isValid && !isPlaceholder) {
      console.log(colorize(`✅ ${varName}`, 'green'));
      console.log(`   ${config.description}`);
      if (varName === 'CLIENT_SECRET') {
        console.log(colorize('   ✓ Real client secret detected', 'green'));
      }
    } else if (hasValue && (isPlaceholder || !isValid)) {
      console.log(colorize(`⚠️  ${varName}`, 'yellow'));
      console.log(`   ${config.description}`);
      if (isPlaceholder) {
        console.log(colorize('   Warning: Still using placeholder value', 'yellow'));
      } else if (!isValid) {
        console.log(colorize('   Warning: Invalid format', 'yellow'));
      }
      warnings.push(varName);
    } else {
      console.log(colorize(`ℹ️  ${varName} (optional)`, 'blue'));
      console.log(`   ${config.description}`);
      if (varName === 'CLIENT_SECRET') {
        console.log(colorize('   Note: Required only for backend token exchange', 'blue'));
      }
    }
    
    console.log(''); // Empty line for readability
  }
  
  // Summary
  console.log(colorize('📊 Validation Summary:', 'cyan'));
  
  if (configIsValid && warnings.length === 0) {
    console.log(colorize('✅ All required environment variables are properly configured!', 'green'));
    console.log(colorize('🚀 Your application should be ready to run.', 'green'));
  } else if (configIsValid && warnings.length > 0) {
    console.log(colorize('⚠️  Configuration is functional but has placeholder values:', 'yellow'));
    warnings.forEach(varName => {
      console.log(colorize(`   - ${varName}`, 'yellow'));
    });
    console.log(colorize('\n💡 Update these with real values for production use.', 'yellow'));
  } else {
    console.log(colorize('❌ Configuration validation failed!', 'red'));
    console.log(colorize('\n🔧 Please fix the above issues and run validation again.', 'red'));
  }
  
  console.log(colorize('\n📚 For detailed setup instructions, see: ENVIRONMENT-SETUP.md\n', 'blue'));
  
  return configIsValid;
}

// Run validation if this script is executed directly
if (require.main === module) {
  const isValid = validateEnvFile();
  process.exit(isValid ? 0 : 1);
}

module.exports = { validateEnvFile };
