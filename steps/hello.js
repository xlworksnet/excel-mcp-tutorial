// hello.js — 연결 확인용으로 만드는 가장 작은 MCP 서버 (5회)
import { McpServer } from '@modelcontextprotocol/server';
import { serveStdio } from '@modelcontextprotocol/server/stdio';
import { z } from 'zod';

function createServer() {
  const server = new McpServer({ name: 'hello', version: '1.0.0' });

  server.registerTool(
    'say_hello',
    {
      description: '이름을 받아 인사말을 돌려줍니다.',
      inputSchema: z.object({ name: z.string().describe('인사할 사람 이름') })
    },
    async ({ name }) => ({
      content: [{ type: 'text', text: `${name}님, 안녕하세요! MCP 서버가 잘 연결되었습니다.` }]
    })
  );

  return server;
}

serveStdio(createServer);
console.error('hello 서버가 시작되었습니다.');   // 확인용 출력은 console.error로!
