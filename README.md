# MultiTabLauncher

A lightweight application launcher that organizes programs into customizable tabs with configurable button layouts.

## Features

- **One-click program execution** - Launch applications instantly with a single button click
- **Customizable button configuration** - Set program path, display name, and administrator privileges for each button
- **Flexible tab management** - Adjust the number of tabs to organize your applications
- **Configurable grid layout** - Control the number of rows and columns of buttons per tab
- **Ultra-lightweight** - Application size is less than 300KB

## Configuration

### Quick Setup
Right-click on any button to configure its settings through the context menu.

### Advanced Configuration
For detailed customization, edit the `MultiTabLauncher.ini` file located in the same folder as `MultiTabLauncher.exe` using any text editor.

**Important**: For multilingual support, save the configuration file in UTF-16 LE BOM encoding.

### Configuration Format

```ini
[Tabs]
Count=10   
ButtonRows=3
ButtonCols=8
Tab0=Home
Tab1=System
```

**Parameters:**
- `Count` - Number of tabs (maximum: 50)
- `ButtonRows` - Number of button rows per tab
- `ButtonCols` - Number of button columns per tab  
- `Tab0`, `Tab1`, etc. - Names for each tab

### Auto-Configuration
If `MultiTabLauncher.ini` doesn't exist when launching the program, it will be automatically created with default settings.

## Development Environment

- **IDE**: Visual Studio 2022
- **Platform**: Win32 API
- **Toolset**: Visual Studio 2022 (v143)
- **C++ Standard**: ISO C++20 (/std:c++20)
- **C Standard**: ISO C17
- **Windows SDK**: 10.0.26100.0

## Getting Started

1. Download and extract the application
2. Run `MultiTabLauncher.exe`
3. Right-click on buttons to configure your applications
4. Customize tabs and layout by editing the INI file as needed

Perfect for users who want a clean, efficient way to organize and launch their frequently used applications without the overhead of complex launcher software.
