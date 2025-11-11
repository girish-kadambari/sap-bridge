# SAP Bridge - C# COM Interop Service

**A high-performance C# service that bridges SAP GUI COM objects with REST APIs**

---

## 🎯 Overview

SAP Bridge is a Windows-based ASP.NET Core service that exposes SAP GUI Scripting API via REST endpoints. It handles COM interop complexity, providing a clean HTTP API for any language to automate SAP.

### Key Features

- ✅ **Native COM Interop** - Direct access to SAP GUI COM objects
- ✅ **REST API** - 8 endpoints for complete SAP control
- ✅ **Dynamic Introspection** - Discover object methods/properties at runtime
- ✅ **Type Safety** - Strong typing with C# models
- ✅ **Session Management** - Multiple concurrent SAP sessions
- ✅ **Error Handling** - Comprehensive exception handling
- ✅ **Cross-Platform Clients** - Use from Mac/Linux via HTTP

---

## 📋 Prerequisites

### Required:
- **Windows OS** (COM objects require Windows)
- **.NET 8.0 SDK** or later
- **SAP GUI for Windows** (installed and licensed)
- **SAP GUI Scripting** enabled

### Enable SAP GUI Scripting:
1. Open SAP GUI
2. Go to: `Options` → `Accessibility & Scripting` → `Scripting`
3. Check: ✅ Enable scripting
4. Check: ✅ Open/Close message when GUI scripting is enabled/disabled

---

## 🚀 Quick Start

### 1. Clone Repository

```bash
git clone https://github.com/yourusername/sap-bridge.git
cd sap-bridge
```

### 2. Build

```bash
dotnet restore
dotnet build
```

### 3. Run

```bash
dotnet run --project SapBridge.Api
```

The service will start on `http://localhost:5000`

### 4. Test

```bash
curl http://localhost:5000/api/health
# Expected: {"status":"healthy","timestamp":"..."}
```

---

## 📡 API Endpoints

### Session Management

**Connect to SAP**
```http
POST /api/session/connect
Content-Type: application/json

{
  "server": null,
  "systemNumber": "00",
  "client": "100"
}

Response: 200 OK
{
  "sessionId": "guid",
  "systemName": "DEV",
  "client": "100",
  "user": "USERNAME"
}
```

**Disconnect**
```http
DELETE /api/session/{sessionId}

Response: 204 No Content
```

### Screen Operations

**Get Screen State**
```http
GET /api/session/{sessionId}/screen

Response: 200 OK
{
  "transaction": "MM03",
  "screenNumber": "100",
  "objects": [...],
  "statusBar": {
    "text": "Material displayed",
    "type": "success"
  }
}
```

### Object Operations

**Get Object Info**
```http
GET /api/object/{sessionId}/info?path=wnd[0]/usr/txtMATNR

Response: 200 OK
{
  "path": "wnd[0]/usr/txtMATNR",
  "type": "GuiTextField",
  "name": "MATNR",
  "label": "Material",
  "methods": ["SetText", "GetText", ...],
  "properties": {...}
}
```

**Invoke Method**
```http
POST /api/invoke
Content-Type: application/json

{
  "sessionId": "guid",
  "objectPath": "wnd[0]/usr/txtMATNR",
  "method": "SetText",
  "args": ["100-400"]
}

Response: 200 OK
{
  "success": true,
  "result": null
}
```

**Find Object**
```http
POST /api/object/{sessionId}/find
Content-Type: application/json

{
  "type": "GuiTextField",
  "name": "MATNR",
  "label": "Material"
}

Response: 200 OK
{
  "found": true,
  "path": "wnd[0]/usr/txtMATNR"
}
```

### Health Check

```http
GET /api/health
GET /api/version
```

---

## 🏗️ Architecture

```
SapBridge/
├── SapBridge.Api/           # ASP.NET Core Web API
│   ├── Controllers/         # REST endpoints
│   │   ├── SessionController.cs
│   │   ├── ScreenController.cs
│   │   ├── ObjectController.cs
│   │   └── HealthController.cs
│   ├── Program.cs
│   └── appsettings.json
│
└── SapBridge.Core/          # Business logic
    ├── Models/              # Data models
    │   ├── ScreenState.cs
    │   ├── ObjectInfo.cs
    │   └── ActionRequest.cs
    ├── SapGuiConnector.cs   # SAP connection mgmt
    ├── ScreenCapture.cs     # Screen state capture
    ├── ActionExecutor.cs    # Method invocation
    └── ComIntrospector.cs   # Dynamic discovery
```

---

## 🔧 Configuration

### appsettings.json

```json
{
  "SapBridge": {
    "SessionTimeout": 3600,
    "MaxSessions": 10,
    "EnableIntrospection": true
  },
  "Logging": {
    "LogLevel": {
      "Default": "Information",
      "SapBridge": "Debug"
    }
  }
}
```

### Environment Variables

```bash
export SAP_BRIDGE_PORT=5000
export SAP_BRIDGE_MAX_SESSIONS=10
```

---

## 🧪 Testing

### Manual Testing

```bash
# 1. Start the service
dotnet run --project SapBridge.Api

# 2. Connect to SAP (another terminal)
curl -X POST http://localhost:5000/api/session/connect \
  -H "Content-Type: application/json" \
  -d '{"client":"100"}'

# Save the sessionId from response

# 3. Get screen state
curl http://localhost:5000/api/session/{sessionId}/screen

# 4. Invoke action
curl -X POST http://localhost:5000/api/invoke \
  -H "Content-Type: application/json" \
  -d '{
    "sessionId": "...",
    "objectPath": "wnd[0]/tbar[0]/okcd",
    "method": "SetText",
    "args": ["MM03"]
  }'
```

### Unit Tests (Coming Soon)

```bash
dotnet test
```

---

## 🌐 Remote Access

### From Mac/Linux

```bash
# Find Windows IP
ipconfig  # Windows: look for IPv4 Address

# Connect from Mac
export SAP_BRIDGE_URL=http://192.168.1.100:5000
curl $SAP_BRIDGE_URL/api/health
```

### Firewall Configuration

```powershell
# Windows Firewall
New-NetFirewallRule -DisplayName "SAP Bridge" `
  -Direction Inbound -LocalPort 5000 -Protocol TCP -Action Allow
```

---

## 📊 Performance

| Metric | Value |
|--------|-------|
| Connection Time | ~2 seconds |
| Screen Capture | <500ms |
| Method Invocation | <100ms |
| Introspection | <200ms |
| Max Concurrent Sessions | 10 |

---

## 🛠️ Development

### Prerequisites

- Visual Studio 2022 or VS Code
- .NET 8.0 SDK
- SAP GUI for Windows

### Build from Source

```bash
git clone https://github.com/yourusername/sap-bridge.git
cd sap-bridge

# Restore dependencies
dotnet restore

# Build
dotnet build

# Run
dotnet run --project SapBridge.Api

# Watch mode (auto-reload)
dotnet watch --project SapBridge.Api
```

### Code Structure

```csharp
// Example: Add new endpoint
[ApiController]
[Route("api/[controller]")]
public class CustomController : ControllerBase
{
    private readonly ISapGuiConnector _connector;
    
    public CustomController(ISapGuiConnector connector)
    {
        _connector = connector;
    }
    
    [HttpPost("custom-action")]
    public IActionResult CustomAction([FromBody] CustomRequest request)
    {
        // Your logic here
    }
}
```

---

## 🐛 Troubleshooting

### SAP GUI Scripting Not Enabled

**Error:** `SAP GUI Scripting is not enabled`

**Solution:**
1. Open SAP GUI
2. Enable scripting in settings
3. Restart SAP GUI

### COM Object Access Denied

**Error:** `Access denied to COM object`

**Solution:**
- Run as Administrator (first time)
- Register COM objects: `regsvr32 sapfewse.ocx`

### Port Already in Use

**Error:** `Address already in use`

**Solution:**
```bash
# Change port in appsettings.json or:
dotnet run --project SapBridge.Api --urls "http://localhost:5001"
```

### Session Timeout

**Error:** `Session not found`

**Solution:**
- Increase `SessionTimeout` in appsettings.json
- Implement session refresh

---

## 📝 License

MIT License - see LICENSE file for details

---

## 🤝 Contributing

Contributions welcome!

1. Fork the repository
2. Create feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit changes (`git commit -m 'Add AmazingFeature'`)
4. Push to branch (`git push origin feature/AmazingFeature`)
5. Open Pull Request

---

## 🔗 Related Projects

- **[sap-use](https://github.com/yourusername/sap-use)** - Python AI agent that uses this bridge
- **SAP GUI Scripting Docs** - [Official Documentation](https://help.sap.com/docs/sap_gui_for_windows/b47d018c3b9b45e897faf66a6c0885a8/a2e9357389334dc89eecc1fb13999ee3.html)

---

## 📧 Support

- Issues: [GitHub Issues](https://github.com/yourusername/sap-bridge/issues)
- Documentation: [Wiki](https://github.com/yourusername/sap-bridge/wiki)
- Community: [Discussions](https://github.com/yourusername/sap-bridge/discussions)

---

**Built with ❤️ for SAP automation**
