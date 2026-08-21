// MCP Client untuk berkomunikasi dengan excel-mcp-server
// Mengelola proses MCP server dan mengirim JSON-RPC requests
const { spawn } = require('child_process');

class McpClient {
    constructor(command, args) {
        this.command = command;
        this.args = args;
        this.process = null;
        this.requestId = 0;
        this.pendingRequests = new Map();
        this.buffer = '';
        this.initialized = false;
    }

    // Mulai MCP server process
    start() {
        if (this.process) return;
        
        this.process = spawn(this.command, this.args, {
            stdio: ['pipe', 'pipe', 'pipe']
        });

        this.process.stdout.on('data', (data) => {
            this.buffer += data.toString();
            this._processBuffer();
        });

        this.process.stderr.on('data', (data) => {
            // Log stderr tapi jangan fail - pydantic warnings normal
            const msg = data.toString();
            if (!msg.includes('IncompleteFieldDefinitionWarning')) {
                console.error('[MCP STDERR]', msg);
            }
        });

        this.process.on('exit', (code) => {
            console.log(`[MCP] Process exited with code ${code}`);
            this.process = null;
            this.initialized = false;
            // Reject all pending requests
            for (const [id, { reject }] of this.pendingRequests) {
                reject(new Error('MCP server process exited'));
            }
            this.pendingRequests.clear();
        });
    }

    // Process buffer untuk extract JSON-RPC messages
    _processBuffer() {
        const lines = this.buffer.split('\n');
        this.buffer = lines.pop(); // Keep incomplete line
        
        for (const line of lines) {
            if (line.trim()) {
                try {
                    const message = JSON.parse(line);
                    if (message.id !== undefined && this.pendingRequests.has(message.id)) {
                        const { resolve } = this.pendingRequests.get(message.id);
                        this.pendingRequests.delete(message.id);
                        resolve(message);
                    }
                } catch (e) {
                    // Not a JSON message, ignore
                }
            }
        }
    }

    // Kirim JSON-RPC request
    _sendRequest(method, params = {}, timeoutMs = 30000) {
        return new Promise((resolve, reject) => {
            if (!this.process) {
                reject(new Error('MCP server not started'));
                return;
            }
            
            const id = ++this.requestId;
            const request = {
                jsonrpc: '2.0',
                id,
                method,
                params
            };
            
            this.pendingRequests.set(id, { resolve, reject });
            this.process.stdin.write(JSON.stringify(request) + '\n');
            
            // Timeout
            setTimeout(() => {
                if (this.pendingRequests.has(id)) {
                    this.pendingRequests.delete(id);
                    reject(new Error(`Timeout waiting for response to ${method}`));
                }
            }, timeoutMs);
        });
    }

    // Initialize MCP server
    async initialize() {
        if (this.initialized) return;
        
        this.start();
        
        const response = await this._sendRequest('initialize', {
            protocolVersion: '2024-11-05',
            capabilities: {},
            clientInfo: { name: 'excel-agent', version: '1.0.0' }
        });
        
        if (response.error) {
            throw new Error(`MCP initialize failed: ${response.error.message}`);
        }
        
        // Send initialized notification
        this.process.stdin.write(JSON.stringify({
            jsonrpc: '2.0',
            method: 'notifications/initialized'
        }) + '\n');
        
        this.initialized = true;
    }

    // Panggil MCP tool
    async callTool(toolName, args) {
        if (!this.initialized) {
            await this.initialize();
        }
        
        const response = await this._sendRequest('tools/call', {
            name: toolName,
            arguments: args
        });
        
        if (response.error) {
            throw new Error(`MCP tool call failed: ${response.error.message}`);
        }
        
        // Extract result content
        if (response.result && response.result.content) {
            const textContent = response.result.content.find(c => c.type === 'text');
            if (textContent) {
                try {
                    return JSON.parse(textContent.text);
                } catch (e) {
                    return textContent.text;
                }
            }
        }
        
        return response.result;
    }

    // List available tools
    async listTools() {
        if (!this.initialized) {
            await this.initialize();
        }
        
        const response = await this._sendRequest('tools/list', {});
        
        if (response.error) {
            throw new Error(`MCP tools/list failed: ${response.error.message}`);
        }
        
        return response.result.tools;
    }

    // Stop MCP server
    stop() {
        if (this.process) {
            this.process.kill();
            this.process = null;
            this.initialized = false;
        }
    }
}

// Singleton instance untuk excel-mcp-server
let excelMcpClient = null;

function getExcelMcpClient() {
    if (!excelMcpClient) {
        excelMcpClient = new McpClient('uvx', ['excel-mcp-server', 'stdio']);
    }
    return excelMcpClient;
}

module.exports = { McpClient, getExcelMcpClient };
