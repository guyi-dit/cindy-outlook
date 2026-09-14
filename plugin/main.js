/* global cindy */

var SECRET_KEY = 'outlook_account';
var PLUGIN_NAME = 'Outlook';
var BASE = 'https://graph.microsoft.com/v1.0/me';
var SELECT_LIST = 'id,subject,from,toRecipients,receivedDateTime,bodyPreview,isRead,hasAttachments,webLink,parentFolderId';
var SELECT_READ = 'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,body,isRead,hasAttachments,webLink,parentFolderId,isDraft';

function fail(message) {
  return { ok: false, message: message };
}

function clampInt(value, fallback, max) {
  var n = typeof value === 'number' && isFinite(value) ? Math.floor(value) : fallback;
  return Math.min(max, Math.max(1, n));
}

function qs(params) {
  var parts = [];
  Object.keys(params).forEach(function (key) {
    var value = params[key];
    if (value === undefined || value === null || value === '') return;
    parts.push(encodeURIComponent(key) + '=' + encodeURIComponent(String(value)));
  });
  return parts.join('&');
}

function messagesUrl(folder) {
  if (!folder) return BASE + '/messages';
  return BASE + '/mailFolders/' + encodeURIComponent(String(folder).trim()) + '/messages';
}


var FOLDER_SELECT = 'id,displayName,totalItemCount,unreadItemCount,childFolderCount,parentFolderId';
var WELL_KNOWN = {
  inbox: 'inbox',
  drafts: 'drafts',
  sentitems: 'sentitems',
  deleteditems: 'deleteditems',
  junkemail: 'junkemail',
  archive: 'archive',
  outbox: 'outbox',
};

function looksLikeGraphId(value) {
  return /^[A-Za-z0-9+/=_-]{20,}$/.test(value);
}

async function listFolderPage(url, account, callId) {
  var items = [];
  var next = url;
  var guard = 0;
  while (next && guard < 20) {
    guard += 1;
    var res = await api({ url: next, account: account, callId: callId });
    if (res.err) return res;
    var page = (res.data && res.data.value) || [];
    for (var i = 0; i < page.length; i++) items.push(page[i]);
    next = res.data && res.data['@odata.nextLink'] ? res.data['@odata.nextLink'] : '';
  }
  return { items: items };
}

async function collectFolders(account, callId) {
  var root = await listFolderPage(
    BASE + '/mailFolders?' + qs({ $top: 50, $select: FOLDER_SELECT }),
    account,
    callId
  );
  if (root.err) return root;
  var out = [];
  async function walk(folder, parentPath) {
    var name = folder.displayName || '';
    var path = parentPath ? parentPath + '/' + name : name;
    out.push({
      id: folder.id,
      name: name,
      path: path,
      parent_id: folder.parentFolderId || '',
      total: folder.totalItemCount,
      unread: folder.unreadItemCount,
      child_folders: folder.childFolderCount || 0,
    });
    if ((folder.childFolderCount || 0) > 0 && out.length < 400) {
      var kids = await listFolderPage(
        BASE + '/mailFolders/' + encodeURIComponent(folder.id) + '/childFolders?' +
          qs({ $top: 50, $select: FOLDER_SELECT }),
        account,
        callId
      );
      if (kids.err) return kids;
      for (var i = 0; i < kids.items.length; i++) {
        var walked = await walk(kids.items[i], path);
        if (walked && walked.err) return walked;
      }
    }
    return null;
  }
  for (var i = 0; i < root.items.length; i++) {
    var err = await walk(root.items[i], '');
    if (err && err.err) return err;
  }
  return { folders: out };
}

async function resolveFolder(spec, account, callId) {
  var raw = String(spec || '').trim();
  if (!raw) return { err: 'folder 不能为空' };
  var key = raw.toLowerCase();
  if (WELL_KNOWN[key]) {
    return { id: WELL_KNOWN[key], name: WELL_KNOWN[key], path: WELL_KNOWN[key] };
  }
  if (looksLikeGraphId(raw)) {
    return { id: raw, name: raw, path: raw };
  }
  var listed = await collectFolders(account, callId);
  if (listed.err) return listed;
  var folders = listed.folders || [];
  function candidates(pred) {
    return folders.filter(pred);
  }
  var exactId = candidates(function (f) { return f.id === raw; });
  if (exactId.length === 1) return exactId[0];
  var byPath = candidates(function (f) {
    return f.path === raw || (f.path && f.path.toLowerCase() === key);
  });
  if (byPath.length === 1) return byPath[0];
  var byName = candidates(function (f) { return f.name === raw; });
  if (byName.length === 1) return byName[0];
  if (byName.length > 1) {
    return {
      err: '文件夹名「' + raw + '」有多个匹配，请改用 path：' +
        byName.map(function (f) { return f.path; }).join('、'),
    };
  }
  var contains = candidates(function (f) {
    return (f.name && f.name.indexOf(raw) !== -1) || (f.path && f.path.indexOf(raw) !== -1);
  });
  if (contains.length === 1) return contains[0];
  if (contains.length > 1) {
    return {
      err: '文件夹「' + raw + '」不唯一，请改用 path：' +
        contains.map(function (f) { return f.path; }).join('、'),
    };
  }
  return { err: '找不到文件夹「' + raw + '」。先 list_folders 查看 path / id' };
}

async function api(opts) {
  var request = {
    url: opts.url,
    method: opts.method || 'GET',
    headers: { Accept: 'application/json' },
    callId: opts.callId,
  };
  if (opts.preferText) request.headers.Prefer = 'outlook.body-content-type="text"';
  if (opts.account) request.authAccount = opts.account;
  if (opts.body !== undefined) {
    request.headers['Content-Type'] = 'application/json';
    request.body = JSON.stringify(opts.body);
  }
  var response = await cindy.fetch(request);
  if (!response.ok) return { err: response.message };
  var data = null;
  if (response.body) {
    try {
      data = JSON.parse(response.body);
    } catch (_err) {
      return { err: 'Microsoft 返回了无法解析的响应(HTTP ' + response.status + ')' };
    }
  }
  if (response.status < 200 || response.status >= 300) {
    var message = data && data.error && data.error.message
      ? data.error.message
      : (response.body || '').slice(0, 200);
    return { err: 'Outlook API 返回 HTTP ' + response.status + ':' + message };
  }
  return { data: data };
}

async function listAccounts() {
  var response = await fetch('/oauth');
  if (!response.ok) return fail('账号状态查询失败(' + response.status + ')');
  var list = await response.json();
  var entry = list.find(function (item) { return item && item.key === SECRET_KEY; });
  if (!entry || !entry.clientConfigured) {
    return fail('尚未配置 Azure 应用。请用 Hotmail 在 Azure 注册公共客户端，运行 scripts/set-client-id.py 写入 Application (client) ID，然后在 Cindy 插件市场刷新本插件');
  }
  if (!entry.accounts.length) {
    return fail('尚未连接 Outlook 账号，请到「' + PLUGIN_NAME + '」详情页单独授权');
  }
  return {
    ok: true,
    result: {
      accounts: entry.accounts.map(function (account) {
        return {
          id: account.id,
          email: account.label,
          status: account.status,
          is_default: account.isDefault,
        };
      }),
    },
  };
}

function addressOf(person) {
  if (!person || !person.emailAddress) return '';
  var name = person.emailAddress.name || '';
  var address = person.emailAddress.address || '';
  if (name && address) return name + ' <' + address + '>';
  return address || name;
}

function addressList(people) {
  if (!Array.isArray(people)) return '';
  return people.map(addressOf).filter(Boolean).join(', ');
}

function stripHtml(html) {
  return String(html)
    .replace(/<\s*br\s*\/?\s*>/gi, '\n')
    .replace(/<\s*\/\s*p\s*>/gi, '\n')
    .replace(/<\s*\/\s*div\s*>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/g, "'")
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]{2,}/g, ' ')
    .trim();
}

function extractBody(message) {
  var body = message && message.body ? message.body : null;
  if (!body || !body.content) return '';
  var text = String(body.content);
  if (String(body.contentType || '').toLowerCase() === 'html') text = stripHtml(text);
  return text.length > 20000 ? text.slice(0, 20000) + '\n…(正文过长已截断)' : text;
}

function summarize(message) {
  return {
    id: message.id,
    from: addressOf(message.from),
    to: addressList(message.toRecipients),
    subject: message.subject || '',
    date: message.receivedDateTime || '',
    snippet: (message.bodyPreview || '').trim(),
    is_read: !!message.isRead,
    has_attachments: !!message.hasAttachments,
    web_link: message.webLink || '',
  };
}

function parseRecipients(value, field) {
  if (value === undefined || value === null || String(value).trim() === '') {
    return field === 'to' ? { err: 'to 不能为空' } : { ok: [] };
  }
  if (/[\r\n]/.test(String(value))) return { err: field + ' 不得包含换行符' };
  var parts = String(value).split(/[,;]+/);
  var result = [];
  for (var i = 0; i < parts.length; i++) {
    var raw = parts[i].trim();
    if (!raw) continue;
    var name = '';
    var email = raw;
    var matched = raw.match(/^(.*)<([^>]+)>$/);
    if (matched) {
      name = matched[1].trim().replace(/^["']|["']$/g, '');
      email = matched[2].trim();
    }
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
      return { err: field + ' 包含无效邮箱: ' + email };
    }
    var item = { emailAddress: { address: email } };
    if (name) item.emailAddress.name = name;
    result.push(item);
  }
  if (field === 'to' && !result.length) return { err: 'to 不能为空' };
  return { ok: result };
}

function buildMessage(args) {
  if (/[\r\n]/.test(String(args.subject || ''))) return { err: 'subject 不得包含换行符' };
  var to = parseRecipients(args.to, 'to');
  if (to.err) return to;
  var cc = parseRecipients(args.cc, 'cc');
  if (cc.err) return cc;
  var bcc = parseRecipients(args.bcc, 'bcc');
  if (bcc.err) return bcc;
  var message = {
    subject: String(args.subject),
    body: { contentType: 'Text', content: String(args.body_text) },
    toRecipients: to.ok,
  };
  if (cc.ok.length) message.ccRecipients = cc.ok;
  if (bcc.ok.length) message.bccRecipients = bcc.ok;
  return { ok: message };
}

async function outlook(args, callId) {
  var account = args.account;
  if (args.action === 'search') {
    var top = clampInt(args.max_results, 5, 10);
    var query = args.query ? String(args.query).trim() : '';
    var folderId;
    if (args.folder) {
      var resolvedSearch = await resolveFolder(args.folder, account, callId);
      if (resolvedSearch.err) return fail(resolvedSearch.err);
      folderId = resolvedSearch.id;
    } else if (!query) {
      folderId = 'inbox';
    }
    var url;
    if (query) {
      var kql = query.replace(/"/g, '').trim();
      if (!kql) return fail('search 的 query 不能为空');
      url = messagesUrl(folderId) + '?' + qs({
        $search: '"' + kql + '"',
        $top: top,
        $select: SELECT_LIST,
      });
    } else {
      url = messagesUrl(folderId) + '?' + qs({
        $top: top,
        $orderby: 'receivedDateTime desc',
        $select: SELECT_LIST,
      });
    }
    var listed = await api({ url: url, account: account, callId: callId });
    if (listed.err) return fail(listed.err);
    var messages = ((listed.data && listed.data.value) || []).map(summarize);
    return {
      ok: true,
      result: {
        count: messages.length,
        next_link: (listed.data && listed.data['@odata.nextLink']) ? true : false,
        messages: messages,
      },
    };
  }

  if (args.action === 'read') {
    if (!args.message_id) return fail('read 需要 message_id');
    var full = await api({
      url: BASE + '/messages/' + encodeURIComponent(args.message_id) + '?' + qs({ $select: SELECT_READ }),
      account: account,
      callId: callId,
      preferText: true,
    });
    if (full.err) return fail(full.err);
    var msg = full.data;
    return {
      ok: true,
      result: {
        id: msg.id,
        from: addressOf(msg.from),
        to: addressList(msg.toRecipients),
        cc: addressList(msg.ccRecipients),
        subject: msg.subject || '',
        date: msg.receivedDateTime || '',
        is_read: !!msg.isRead,
        is_draft: !!msg.isDraft,
        has_attachments: !!msg.hasAttachments,
        web_link: msg.webLink || '',
        body: extractBody(msg),
      },
    };
  }

  if (args.action === 'list_folders') {
    var collected = await collectFolders(account, callId);
    if (collected.err) return fail(collected.err);
    return {
      ok: true,
      result: {
        folders: collected.folders || [],
        well_known: ['inbox', 'drafts', 'sentitems', 'deleteditems', 'junkemail', 'archive'],
      },
    };
  }

  if (args.action === 'delete') {
    if (!args.message_id) return fail('delete 需要 message_id');
    var removed = await api({
      url: BASE + '/messages/' + encodeURIComponent(args.message_id),
      method: 'DELETE',
      account: account,
      callId: callId,
    });
    if (removed.err) return fail(removed.err);
    return {
      ok: true,
      result: {
        deleted: true,
        id: args.message_id,
        note: '已删除。若邮件原先不在已删除邮件中，会先进入已删除邮件；已在其中则会永久删除。',
      },
    };
  }

  if (args.action === 'move') {
    if (!args.message_id) return fail('move 需要 message_id');
    var dest = args.folder ? String(args.folder).trim() : '';
    if (!dest) return fail('move 需要 folder（目标文件夹名称、path 或 id，例如 富途 或 收件箱/富途）');
    var resolved = await resolveFolder(dest, account, callId);
    if (resolved.err) return fail(resolved.err);
    var moved = await api({
      url: BASE + '/messages/' + encodeURIComponent(args.message_id) + '/move',
      method: 'POST',
      body: { destinationId: resolved.id },
      account: account,
      callId: callId,
    });
    if (moved.err) return fail(moved.err);
    return {
      ok: true,
      result: {
        moved: true,
        id: moved.data && moved.data.id ? moved.data.id : args.message_id,
        destination: resolved.path || resolved.name || dest,
        destination_id: resolved.id,
        subject: moved.data && moved.data.subject ? moved.data.subject : undefined,
      },
    };
  }

  if (args.action === 'send' || args.action === 'draft') {
    if (!args.to || args.subject === undefined || args.body_text === undefined) {
      return fail(args.action + ' 需要 to / subject / body_text');
    }
    var built = buildMessage(args);
    if (built.err) return fail(built.err);
    if (args.action === 'send') {
      var sent = await api({
        url: BASE + '/sendMail',
        method: 'POST',
        body: { message: built.ok, saveToSentItems: true },
        account: account,
        callId: callId,
      });
      if (sent.err) return fail(sent.err);
      return { ok: true, result: { sent: true } };
    }
    var draft = await api({
      url: BASE + '/messages',
      method: 'POST',
      body: built.ok,
      account: account,
      callId: callId,
    });
    if (draft.err) return fail(draft.err);
    return {
      ok: true,
      result: {
        draft: true,
        id: draft.data && draft.data.id,
        web_link: draft.data && draft.data.webLink,
      },
    };
  }

  return fail('未知 action:' + args.action);
}

cindy.onHostMessage(async function (message) {
  if (!message || message.type !== 'tool-call') return;
  try {
    var result = message.tool === 'outlook_accounts'
      ? await listAccounts()
      : message.tool === 'outlook'
        ? await outlook(message.args || {}, message.callId)
        : fail('未知工具:' + message.tool);
    if (result.ok) {
      cindy.send({ type: 'tool-result', callId: message.callId, ok: true, result: result.result });
    } else {
      cindy.send({ type: 'tool-result', callId: message.callId, ok: false, message: result.message });
    }
  } catch (error) {
    cindy.send({
      type: 'tool-result',
      callId: message.callId,
      ok: false,
      message: 'Outlook 工具执行失败:' + (error && error.message ? error.message : String(error)),
    });
  }
});
