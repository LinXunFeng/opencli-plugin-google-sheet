import test from 'node:test';
import assert from 'node:assert/strict';

import {
  buildCookieHeader,
  fetchSheetRows,
  fetchWorkbookMeta,
} from '../src/sheets-api.js';

test('buildCookieHeader joins cookie name/value pairs', () => {
  const value = buildCookieHeader([
    { name: 'A', value: '1' },
    { name: 'B', value: '2' },
  ]);
  assert.equal(value, 'A=1; B=2');
});

test('fetchWorkbookMeta merges cookies from url, docs domain and parent google domain', async () => {
  const calls = [];
  const page = {
    getCookies: async (input) => {
      calls.push(input);
      if (input?.url) {
        return [{ name: 'sid', value: 'u', domain: 'docs.google.com' }];
      }
      if (input?.domain === 'google.com') {
        return [{ name: 'ssid', value: 'g', domain: '.google.com' }];
      }
      return [{ name: 'sap', value: 'd', domain: 'docs.google.com' }];
    },
  };

  globalThis.fetch = async (url, options) => {
    assert.match(String(url), /spreadsheets\/d\/doc123\/edit/);
    assert.equal(options.headers.Cookie, 'sap=d; ssid=g; sid=u');
    return {
      ok: true,
      status: 200,
      url: String(url),
      text: async () => '{"sheetId":0,"title":"Sheet1"}{"sheetId":1,"title":"Sheet2"}',
    };
  };

  const sheets = await fetchWorkbookMeta(page, 'doc123');
  assert.deepEqual(sheets, [
    { gid: '0', title: 'Sheet1', index: 0 },
    { gid: '1', title: 'Sheet2', index: 1 },
  ]);
  assert.deepEqual(calls, [
    { url: 'https://docs.google.com/spreadsheets/d/doc123/edit' },
    { domain: 'docs.google.com' },
    { domain: 'google.com' },
  ]);
});

test('fetchWorkbookMeta does not treat embedded ServiceLogin links as unauthenticated when sheet metadata exists', async () => {
  const page = {
    getCookies: async () => [],
  };

  globalThis.fetch = async () => ({
    ok: true,
    status: 200,
    url: 'https://docs.google.com/spreadsheets/d/doc123/edit',
    text: async () => '<a href="https://accounts.google.com/ServiceLogin">switch account</a>{"sheetId":0,"title":"Sheet1"}',
  });

  const sheets = await fetchWorkbookMeta(page, 'doc123');
  assert.deepEqual(sheets, [{ gid: '0', title: 'Sheet1', index: 0 }]);
});

test('fetchWorkbookMeta throws AUTH_REQUIRED on login redirect', async () => {
  const page = {
    getCookies: async () => [],
  };

  globalThis.fetch = async () => ({
    ok: true,
    status: 200,
    url: 'https://accounts.google.com/ServiceLogin',
    text: async () => '<html>login</html>',
  });

  await assert.rejects(() => fetchWorkbookMeta(page, 'doc123'), (error) => {
    assert.equal(error.code, 'AUTH_REQUIRED');
    return true;
  });
});

test('fetchSheetRows parses gviz callback text', async () => {
  const page = {
    getCookies: async () => [{ name: 'sid', value: 'x', domain: 'docs.google.com' }],
  };

  globalThis.fetch = async (url) => {
    assert.match(String(url), /gviz\/tq\?gid=0/);
    return {
      ok: true,
      status: 200,
      text: async () => 'google.visualization.Query.setResponse({"status":"ok","table":{"cols":[{},{}],"rows":[{"c":[{"v":"H1"},{"v":"H2"}]},{"c":[{"v":"A"},{"v":"1"}]}]}});',
    };
  };

  const rows = await fetchSheetRows(page, 'doc123', '0');
  assert.deepEqual(rows, [
    ['H1', 'H2'],
    ['A', '1'],
  ]);
});

test('fetchSheetRows maps gviz access errors to AUTH_REQUIRED', async () => {
  const page = { getCookies: async () => [] };
  globalThis.fetch = async () => ({
    ok: true,
    status: 200,
    text: async () => 'google.visualization.Query.setResponse({"status":"error","errors":[{"reason":"access_denied","message":"Denied"}]});',
  });

  await assert.rejects(() => fetchSheetRows(page, 'doc123', '0'), (error) => {
    assert.equal(error.code, 'AUTH_REQUIRED');
    return true;
  });
});

test('fetchWorkbookMeta prefers browser-context fetch when available', async () => {
  const page = {
    evaluate: async () => ({
      ok: true,
      status: 200,
      url: 'https://docs.google.com/spreadsheets/d/doc123/edit',
      body: '{"sheetId":0,"title":"Sheet1"}',
    }),
    getCookies: async () => {
      throw new Error('should not read cookies when browser fetch succeeds');
    },
  };

  globalThis.fetch = async () => {
    throw new Error('should not call node fetch when browser fetch succeeds');
  };

  const sheets = await fetchWorkbookMeta(page, 'doc123');
  assert.deepEqual(sheets, [{ gid: '0', title: 'Sheet1', index: 0 }]);
});

test('fetchSheetRows prefers browser CSV export when available', async () => {
  const page = {
    evaluate: async (script) => {
      if (script.includes('/export?format=csv')) {
        return {
          ok: true,
          status: 200,
          url: 'https://docs.google.com/spreadsheets/d/doc123/export?format=csv&gid=0',
          body: 'Name,Score\nAlice,90\nBob,80\n',
        };
      }
      throw new Error('unexpected evaluate call');
    },
    getCookies: async () => {
      throw new Error('should not use cookie fetch when CSV export succeeds');
    },
  };

  globalThis.fetch = async () => {
    throw new Error('should not call node fetch when CSV export succeeds');
  };

  const rows = await fetchSheetRows(page, 'doc123', '0');
  assert.deepEqual(rows, [
    ['Name', 'Score'],
    ['Alice', '90'],
    ['Bob', '80'],
  ]);
});

test('fetchSheetRows falls back to visible DOM grid when CSV and gviz are denied', async () => {
  const page = {
    goto: async () => {},
    wait: async () => {},
    evaluate: async (script) => {
      if (script.includes('/export?format=csv')) {
        return {
          ok: true,
          status: 200,
          url: 'https://docs.google.com/spreadsheets/d/doc123/export?format=csv&gid=0',
          body: 'ACCESS_DENIED',
        };
      }
      if (script.includes('[role="gridcell"]')) {
        return [
          ['Name', 'Score'],
          ['Alice', '90'],
        ];
      }
      throw new Error('unexpected evaluate call');
    },
    getCookies: async () => [{ name: 'sid', value: 'x', domain: 'docs.google.com' }],
  };

  globalThis.fetch = async () => ({
    ok: true,
    status: 200,
    url: 'https://docs.google.com/spreadsheets/d/doc123/gviz/tq?gid=0&tqx=out:json',
    text: async () => 'google.visualization.Query.setResponse({"status":"error","errors":[{"reason":"access_denied","message":"ACCESS_DENIED"}]});',
  });

  const rows = await fetchSheetRows(page, 'doc123', '0');
  assert.deepEqual(rows, [
    ['Name', 'Score'],
    ['Alice', '90'],
  ]);
});

test('fetchSheetRows falls back to node cookie csv fetch when context fetch is unavailable', async () => {
  const page = {
    evaluate: async (script) => {
      if (script.includes('/export?format=csv')) {
        // Simulate browser-context fetch unavailable
        return { ok: false, status: 0, url: 'https://docs.google.com/spreadsheets/d/doc123/export?format=csv&gid=0', body: '' };
      }
      throw new Error('unexpected evaluate call');
    },
    getCookies: async (input) => {
      if (input?.url) {
        return [{ name: 'sid', value: 'u', domain: 'docs.google.com' }];
      }
      if (input?.domain === 'google.com') {
        return [{ name: 'ssid', value: 'g', domain: '.google.com' }];
      }
      return [{ name: 'sap', value: 'd', domain: 'docs.google.com' }];
    },
  };

  let called = 0;
  globalThis.fetch = async (url, options) => {
    called += 1;
    if (String(url).includes('/export?format=csv')) {
      assert.match(options.headers.Cookie, /sid=u/);
      return {
        ok: true,
        status: 200,
        url: String(url),
        text: async () => 'Name,Score\nAlice,90\n',
      };
    }
    throw new Error('should not call gviz when node csv succeeds');
  };

  const rows = await fetchSheetRows(page, 'doc123', '0');
  assert.deepEqual(rows, [
    ['Name', 'Score'],
    ['Alice', '90'],
  ]);
  assert.equal(called, 1);
});

test('fetchWorkbookMeta reads sheet metadata from the loaded browser page before fallback', async () => {
  const calls = [];
  const page = {
    goto: async (url) => {
      calls.push(['goto', url]);
    },
    evaluate: async (script) => {
      calls.push(['evaluate', script]);
      if (script === 'location.href') {
        return 'https://docs.google.com/spreadsheets/d/doc123/edit';
      }
      return {
        sheets: [{ gid: '0', title: 'Sheet1', index: 0 }],
        confident: true,
      };
    },
    getCookies: async () => {
      throw new Error('should not use cookie fallback when page extraction succeeds');
    },
  };

  globalThis.fetch = async () => {
    throw new Error('should not call node fetch when page extraction succeeds');
  };

  const sheets = await fetchWorkbookMeta(page, 'doc123');
  assert.deepEqual(sheets, [{ gid: '0', title: 'Sheet1', index: 0 }]);
  assert.equal(calls[0][0], 'goto');
  assert.equal(calls[1][0], 'evaluate');
  assert.equal(calls[2][0], 'evaluate');
});

test('fetchWorkbookMeta retries with session fetch when page extraction is low confidence', async () => {
  const page = {
    goto: async () => {},
    wait: async () => {},
    evaluate: async (script) => {
      if (script === 'location.href') {
        return 'https://docs.google.com/spreadsheets/d/doc123/edit?gid=2118959825';
      }
      if (script.includes('fetch(')) {
        return {
          ok: false,
          status: 0,
          url: 'https://docs.google.com/spreadsheets/d/doc123/edit',
          body: '',
        };
      }
      return {
        sheets: [{ gid: '2118959825', title: 'Doc Title - Google 表格', index: 0 }],
        confident: false,
      };
    },
    getCookies: async () => [],
  };

  globalThis.fetch = async () => ({
    ok: true,
    status: 200,
    url: 'https://docs.google.com/spreadsheets/d/doc123/edit',
    text: async () => '{"sheetId":101,"gridProperties":{"rowCount":1000},"title":"Backlog"}{"sheetId":202,"gridProperties":{"rowCount":1000},"title":"Roadmap"}',
  });

  const sheets = await fetchWorkbookMeta(page, 'doc123');
  assert.deepEqual(sheets, [
    { gid: '101', title: 'Backlog', index: 0 },
    { gid: '202', title: 'Roadmap', index: 1 },
  ]);
});

test('fetchWorkbookMeta falls back to htmlview when edit metadata is incomplete', async () => {
  const page = {
    goto: async () => {},
    wait: async () => {},
    evaluate: async (script) => {
      if (script === 'location.href') {
        return 'https://docs.google.com/spreadsheets/d/doc123/edit?gid=2118959825';
      }
      if (script.includes('fetch(')) {
        return {
          ok: false,
          status: 0,
          url: 'https://docs.google.com/spreadsheets/d/doc123/edit',
          body: '',
        };
      }
      return {
        sheets: [{ gid: '2118959825', title: 'Doc Title - Google 表格', index: 0 }],
        confident: false,
      };
    },
    getCookies: async () => [],
  };

  let callCount = 0;
  globalThis.fetch = async (url) => {
    callCount += 1;
    if (String(url).includes('/edit')) {
      return {
        ok: true,
        status: 200,
        url: String(url),
        text: async () => '{"sheetId":2118959825,"title":"计划表"}',
      };
    }
    if (String(url).includes('/htmlview')) {
      return {
        ok: true,
        status: 200,
        url: String(url),
        text: async () => `
          <a href="/spreadsheets/d/doc123/htmlview?gid=2118959825">计划表</a>
          <a href="/spreadsheets/d/doc123/htmlview?gid=638061341">Roadmap</a>
        `,
      };
    }
    throw new Error(`unexpected url: ${String(url)}`);
  };

  const sheets = await fetchWorkbookMeta(page, 'doc123');
  assert.deepEqual(sheets, [
    { gid: '2118959825', title: '计划表', index: 0 },
    { gid: '638061341', title: 'Roadmap', index: 1 },
  ]);
  assert.equal(callCount >= 2, true);
});

test('fetchWorkbookMeta prefers htmlview metadata with richer titles when gid count ties', async () => {
  const page = {
    goto: async () => {},
    wait: async () => {},
    evaluate: async (script) => {
      if (script === 'location.href') {
        return 'https://docs.google.com/spreadsheets/d/doc123/edit?gid=2118959825';
      }
      if (script.includes('fetch(')) {
        return {
          ok: false,
          status: 0,
          url: 'https://docs.google.com/spreadsheets/d/doc123/edit',
          body: '',
        };
      }
      if (script === 'document.documentElement ? document.documentElement.outerHTML : ""') {
        return '';
      }
      return {
        sheets: [{ gid: '2118959825', title: 'Doc Title - Google 表格', index: 0 }],
        confident: false,
      };
    },
    getCookies: async () => [],
  };

  let htmlViewCalls = 0;
  globalThis.fetch = async (url) => {
    const textUrl = String(url);
    if (textUrl.includes('/edit')) {
      return {
        ok: true,
        status: 200,
        url: textUrl,
        text: async () => '{"sheetId":2118959825,"title":"计划表"}',
      };
    }
    if (textUrl.includes('/htmlview?rm=minimal')) {
      htmlViewCalls += 1;
      return {
        ok: true,
        status: 200,
        url: textUrl,
        text: async () => `
          <script>
            const a = "/spreadsheets/d/doc123/htmlview?rm=minimal&gid=2118959825";
            const b = "/spreadsheets/d/doc123/htmlview?rm=minimal&gid=638061341";
          </script>
        `,
      };
    }
    if (textUrl.includes('/htmlview')) {
      htmlViewCalls += 1;
      return {
        ok: true,
        status: 200,
        url: textUrl,
        text: async () => `
          <a href="/spreadsheets/d/doc123/htmlview?gid=2118959825">计划表</a>
          <a href="/spreadsheets/d/doc123/htmlview?gid=638061341">Roadmap</a>
        `,
      };
    }
    throw new Error(`unexpected url: ${textUrl}`);
  };

  const sheets = await fetchWorkbookMeta(page, 'doc123');
  assert.deepEqual(sheets, [
    { gid: '2118959825', title: '计划表', index: 0 },
    { gid: '638061341', title: 'Roadmap', index: 1 },
  ]);
  assert.equal(htmlViewCalls, 2);
});

test('fetchWorkbookMeta uses htmlview DOM fallback when htmlview fetch is weak', async () => {
  const page = {
    goto: async () => {},
    wait: async () => {},
    evaluate: async (script) => {
      if (script === 'location.href') {
        return 'https://docs.google.com/spreadsheets/d/doc123/edit?gid=2118959825';
      }
      if (script.includes('fetch(')) {
        return {
          ok: false,
          status: 0,
          url: 'https://docs.google.com/spreadsheets/d/doc123/edit',
          body: '',
        };
      }
      if (script === 'document.documentElement ? document.documentElement.outerHTML : ""') {
        return `
          <script>
            const a = "/spreadsheets/d/doc123/htmlview?rm=minimal&gid=2118959825";
            const b = "/spreadsheets/d/doc123/htmlview?rm=minimal&gid=638061341";
          </script>
        `;
      }
      return {
        sheets: [{ gid: '2118959825', title: 'Doc Title - Google 表格', index: 0 }],
        confident: false,
      };
    },
    getCookies: async () => [],
  };

  globalThis.fetch = async (url) => {
    if (String(url).includes('/edit')) {
      return {
        ok: true,
        status: 200,
        url: String(url),
        text: async () => '{"sheetId":2118959825,"title":"计划表"}',
      };
    }
    if (String(url).includes('/htmlview')) {
      return {
        ok: true,
        status: 200,
        url: String(url),
        text: async () => '{"sheetId":2118959825,"title":"计划表"}',
      };
    }
    throw new Error(`unexpected url: ${String(url)}`);
  };

  const sheets = await fetchWorkbookMeta(page, 'doc123');
  assert.deepEqual(sheets, [
    { gid: '2118959825', title: '计划表', index: 0 },
    { gid: '638061341', title: 'Sheet_638061341', index: 1 },
  ]);
});

test('fetchWorkbookMeta hydrates generic htmlview titles from live edit tabs', async () => {
  const page = {
    goto: async () => {},
    wait: async () => {},
    evaluate: async (script) => {
      if (script === 'location.href') {
        return 'https://docs.google.com/spreadsheets/d/doc123/edit?gid=2118959825';
      }
      if (script.includes('fetch(')) {
        return {
          ok: false,
          status: 0,
          url: 'https://docs.google.com/spreadsheets/d/doc123/edit',
          body: '',
        };
      }
      if (script.includes('recover real titles')) {
        return [
          { gid: '2118959825', title: '计划表' },
          { gid: '638061341', title: 'Roadmap' },
        ];
      }
      if (script === 'document.documentElement ? document.documentElement.outerHTML : ""') {
        return `
          <script>
            const a = "/spreadsheets/d/doc123/htmlview?rm=minimal&gid=2118959825";
            const b = "/spreadsheets/d/doc123/htmlview?rm=minimal&gid=638061341";
          </script>
        `;
      }
      return {
        sheets: [{ gid: '2118959825', title: 'Doc Title - Google 表格', index: 0 }],
        confident: false,
      };
    },
    getCookies: async () => [],
  };

  globalThis.fetch = async (url) => {
    const textUrl = String(url);
    if (textUrl.includes('/edit')) {
      return {
        ok: true,
        status: 200,
        url: textUrl,
        text: async () => '{"sheetId":2118959825,"title":"计划表"}',
      };
    }
    if (textUrl.includes('/htmlview')) {
      return {
        ok: true,
        status: 200,
        url: textUrl,
        text: async () => `
          <script>
            const a = "/spreadsheets/d/doc123/htmlview?rm=minimal&gid=2118959825";
            const b = "/spreadsheets/d/doc123/htmlview?rm=minimal&gid=638061341";
          </script>
        `,
      };
    }
    throw new Error(`unexpected url: ${textUrl}`);
  };

  const sheets = await fetchWorkbookMeta(page, 'doc123');
  assert.deepEqual(sheets, [
    { gid: '2118959825', title: '计划表', index: 0 },
    { gid: '638061341', title: 'Roadmap', index: 1 },
  ]);
});
