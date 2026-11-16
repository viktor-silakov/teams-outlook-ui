(function(){const t=document.createElement("link").relList;if(t&&t.supports&&t.supports("modulepreload"))return;for(const n of document.querySelectorAll('link[rel="modulepreload"]'))l(n);new MutationObserver(n=>{for(const o of n)if(o.type==="childList")for(const c of o.addedNodes)c.tagName==="LINK"&&c.rel==="modulepreload"&&l(c)}).observe(document,{childList:!0,subtree:!0});function s(n){const o={};return n.integrity&&(o.integrity=n.integrity),n.referrerPolicy&&(o.referrerPolicy=n.referrerPolicy),n.crossOrigin==="use-credentials"?o.credentials="include":n.crossOrigin==="anonymous"?o.credentials="omit":o.credentials="same-origin",o}function l(n){if(n.ep)return;n.ep=!0;const o=s(n);fetch(n.href,o)}})();const ve="modulepreload",ye=function(a,t){return new URL(a,t).href},Y={},F=function(t,s,l){let n=Promise.resolve();if(s&&s.length>0){const c=document.getElementsByTagName("link"),m=document.querySelector("meta[property=csp-nonce]"),r=m?.nonce||m?.getAttribute("nonce");n=Promise.allSettled(s.map(i=>{if(i=ye(i,l),i in Y)return;Y[i]=!0;const u=i.endsWith(".css"),h=u?'[rel="stylesheet"]':"";if(!!l)for(let I=c.length-1;I>=0;I--){const T=c[I];if(T.href===i&&(!u||T.rel==="stylesheet"))return}else if(document.querySelector(`link[href="${i}"]${h}`))return;const g=document.createElement("link");if(g.rel=u?"stylesheet":ve,u||(g.as="script"),g.crossOrigin="",g.href=i,r&&g.setAttribute("nonce",r),document.head.appendChild(g),u)return new Promise((I,T)=>{g.addEventListener("load",I),g.addEventListener("error",()=>T(new Error(`Unable to preload CSS for ${i}`)))})}))}function o(c){const m=new Event("vite:preloadError",{cancelable:!0});if(m.payload=c,window.dispatchEvent(m),!m.defaultPrevented)throw c}return n.then(c=>{for(const m of c||[])m.status==="rejected"&&o(m.reason);return t().catch(o)})},E=window.location.hostname==="localhost"||window.location.hostname==="127.0.0.1"?"":"https://localhost:3978",D="teams_outlook_sid";function we(){try{const a=localStorage.getItem(D);if(a)return a;const t=sessionStorage.getItem(D);if(t)return localStorage.setItem(D,t),t}catch{}return null}class se extends Error{constructor(t="Local server at localhost:3978 is unreachable"){super(t),this.name="ServerUnreachableError"}}let W=null,Q=we();function $e(a){W=a}function q(a){Q=a;try{a?(localStorage.setItem(D,a),sessionStorage.setItem(D,a)):(localStorage.removeItem(D),sessionStorage.removeItem(D))}catch{}}function Ie(a,t){const s=new Headers(a);return t&&!s.has("Content-Type")&&s.set("Content-Type","application/json"),W?(s.set("Authorization",`Bearer ${W}`),s.set("x-token-type","teams-sso")):Q&&s.set("x-session-id",Q),s}async function ne(a,t={}){const s=Ie(t.headers,!!t.body);let l;try{l=await fetch(`${E}${a}`,{...t,headers:s,credentials:"include"})}catch(n){throw n instanceof DOMException&&n.name==="AbortError"?n:new se}if(!l.ok){let n=`Request failed (${l.status})`;try{const o=await l.json();o.error?.message&&(n=o.error.message)}catch{}throw new Error(n)}return l}async function b(a,t={}){return await(await ne(a,t)).json()}function P(){return b("/api/auth/status")}function ke(){return b("/api/auth/logout",{method:"POST"})}function oe(){return b("/api/mail/folders")}function N(a,t){const s=new URLSearchParams({folder:a.folder,unreadOnly:String(a.unreadOnly),top:String(a.top??25),skip:String(a.skip??0)});return a.search.trim()&&s.set("search",a.search.trim()),b(`/api/mail/messages?${s.toString()}`,{signal:t})}function Ae(a,t){return b(`/api/mail/messages/${encodeURIComponent(a)}`,{signal:t})}function Te(a,t){return b(`/api/mail/conversations/${encodeURIComponent(a)}`,{signal:t})}function Oe(a,t){return b("/api/mail/bulk",{method:"POST",body:JSON.stringify({ids:a,action:t})})}function re(a){return b("/api/mail/drafts",{method:"POST",body:JSON.stringify(a)})}function Ee(a){return b(`/api/mail/drafts/${encodeURIComponent(a)}/send`,{method:"POST"})}function ie(a,t){return b(`/api/mail/messages/${encodeURIComponent(a)}/read`,{method:"PATCH",body:JSON.stringify({isRead:t})})}async function De(a,t,s){const n=await(await ne(`/api/mail/messages/${encodeURIComponent(a)}/attachments/${encodeURIComponent(t)}`)).blob(),o=URL.createObjectURL(n),c=document.createElement("a");c.href=o,c.download=s||"attachment",document.body.appendChild(c),c.click(),c.remove(),setTimeout(()=>URL.revokeObjectURL(o),1e3)}function le(a){return b(`/api/mail/messages/${encodeURIComponent(a)}`,{method:"DELETE"})}function Me(a){return b(`/api/mail/messages/${encodeURIComponent(a)}/archive`,{method:"POST"})}function Se(a,t,s){return b(`/api/mail/messages/${encodeURIComponent(a)}/forward`,{method:"POST",body:JSON.stringify({to:t,comment:s})})}function Ce(a,t,s=!1,l){return b(`/api/mail/messages/${encodeURIComponent(a)}/reply`,{method:"POST",body:JSON.stringify({comment:t,replyAll:s,attachments:l})})}function Be(a){return b("/api/mail/messages/send",{method:"POST",body:JSON.stringify(a)})}function Le(a,t){return b(`/api/mail/recipients?q=${encodeURIComponent(a)}`,{signal:t})}function Re(a=!1){return b(`/api/app/version${a?"?check=1":""}`)}function xe(){return b("/api/app/update",{method:"POST"})}function X(a){return a==="dark"||a==="contrast"?a:"default"}function z(a,t,s){return Promise.race([a,new Promise((l,n)=>setTimeout(()=>n(new Error(`${s} timeout`)),t))])}async function Ue(a){if(!(window.parent!==window||!!window.opener))return{inTeams:!1,theme:"default",ssoToken:null};try{const{app:s,authentication:l}=await F(async()=>{const{app:m,authentication:r}=await import("./src-CZ2ouLw7.js");return{app:m,authentication:r}},[],import.meta.url);await z(s.initialize(),500,"Teams init");const n=await z(s.getContext(),1e3,"Teams getContext"),o=X(n.app?.theme);s.registerOnThemeChangeHandler(m=>{a(X(m))});let c=null;try{c=await z(l.getAuthToken(),1200,"Teams getAuthToken")}catch{c=null}return{inTeams:!0,theme:o,ssoToken:c}}catch{return{inTeams:!1,theme:"default",ssoToken:null}}}const de=[{id:"inbox",label:"Inbox"},{id:"sentitems",label:"Sent"},{id:"drafts",label:"Drafts"},{id:"archive",label:"Archive"},{id:"deleteditems",label:"Deleted"}],A=document.getElementById("root"),O=new Map;let e={theme:"default",inTeams:!1,auth:null,serverUnreachable:!1,folder:"inbox",folders:[],unreadOnly:!1,search:"",messages:[],hasMore:!1,loadingMore:!1,selectedId:null,selectedConversationId:null,selectedDetail:null,selectedThread:[],checkedIds:new Set,loadingList:!1,loadingDetail:!1,replyOpen:!1,replyAll:!1,sendingReply:!1,replyAttachments:[],forwardOpen:!1,sendingForward:!1,composeOpen:!1,composeDraftId:null,composeInitial:{to:"",subject:"",body:""},draftStatus:"",sendingCompose:!1,composeAttachments:[],updatingApp:!1,userMenuOpen:!1,versionInfo:null,error:null};const k=[];let $=-1,M=null,V=null,j=null,K=null,J=null,G=null,v=[],w=-1;function d(a){return a?a.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}function L(a){const t=new Date(a);return Number.isNaN(t.getTime())?a:t.toLocaleString([],{month:"short",day:"numeric",hour:"2-digit",minute:"2-digit"})}function ee(a){e.theme=a,document.documentElement.setAttribute("data-theme",a)}function je(a){return a<1024?`${a} B`:a<1024*1024?`${(a/1024).toFixed(1)} KB`:`${(a/(1024*1024)).toFixed(1)} MB`}async function te(a){if(!a||a.length===0)return[];const t=[];for(const s of Array.from(a).slice(0,10)){const l=await s.arrayBuffer(),n=new Uint8Array(l);let o="";const c=32768;for(let m=0;m<n.length;m+=c)o+=String.fromCharCode(...n.subarray(m,m+c));t.push({name:s.name,contentType:s.type||"application/octet-stream",contentBytes:btoa(o)})}return t}function B(a,t){return a.length===0?"":`
    <div class="attachment-chips">
      ${a.map((s,l)=>`
          <span class="attachment-chip">
            <span>${d(s.name)}</span>
            <button type="button" class="chip-remove" data-remove-att="${t}" data-att-idx="${l}" aria-label="Remove ${d(s.name)}">&times;</button>
          </span>`).join("")}
    </div>`}function ce(a){const t=new Map;for(const s of a){const l=s.conversationId||s.id,n=t.get(l);n?(n.count+=1,s.isRead||(n.hasUnread=!0),n.ids.push(s.id)):t.set(l,{key:l,conversationId:s.conversationId,latest:s,count:1,hasUnread:!s.isRead,ids:[s.id]})}return Array.from(t.values())}const S={markRead:'<svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M2 6l6 4 6-4"/><path d="M2 6v6a1 1 0 001 1h10a1 1 0 001-1V6l-6-3.5L2 6z"/></svg>',markUnread:'<svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="3.5" width="12" height="9" rx="1"/><path d="M2.5 4l5.5 4 5.5-4"/></svg>',archive:'<svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="3" width="12" height="3" rx="0.5"/><path d="M3 6v6.5a1 1 0 001 1h8a1 1 0 001-1V6"/><path d="M6.5 9h3"/></svg>',delete:'<svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M2.5 4h11"/><path d="M6 4V2.5h4V4"/><path d="M4 4l.7 9a1 1 0 001 .9h4.6a1 1 0 001-.9L12 4"/></svg>',clear:'<svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M4 4l8 8M12 4l-8 8"/></svg>',unreadFilter:'<svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M2 4h12l-4.5 5v3.5l-3 1.5V9L2 4z"/></svg>'};function R(){const a=ce(e.messages),t=e.checkedIds.size,s=t>0,l=a.length>0&&a.every(r=>e.checkedIds.has(r.latest.id)),n=s?"":"disabled",o=`
    <div class="bulk-bar" role="toolbar" aria-label="Message actions">
      <div class="bulk-left">
        <input id="select-all-chk" type="checkbox" class="msg-check-all" ${l?"checked":""} ${a.length===0?"disabled":""} title="${l?"Deselect all visible messages":"Select all visible messages"}" aria-label="Select all visible messages" />
        <span class="bulk-count">${s?`${t} selected`:""}</span>
        <button type="button" class="icon-btn" data-bulk="markRead" title="Mark selected as read" aria-label="Mark selected as read" ${n}>${S.markRead}</button>
        <button type="button" class="icon-btn" data-bulk="markUnread" title="Mark selected as unread" aria-label="Mark selected as unread" ${n}>${S.markUnread}</button>
        <button type="button" class="icon-btn" data-bulk="archive" title="Archive selected" aria-label="Archive selected" ${n}>${S.archive}</button>
        <button type="button" class="icon-btn icon-btn-danger" data-bulk="delete" title="Delete selected" aria-label="Delete selected" ${n}>${S.delete}</button>
        <button type="button" class="icon-btn" id="bulk-clear-btn" title="Clear selection" aria-label="Clear selection" ${n}>${S.clear}</button>
      </div>
      <button type="button" id="unread-filter-btn" class="icon-btn unread-toggle-btn ${e.unreadOnly?"active":""}" title="${e.unreadOnly?"Unread only (click to show Read & Unread)":"Filter: Unread only (click to toggle)"}" aria-pressed="${e.unreadOnly}" aria-label="Toggle unread only filter">
        ${S.unreadFilter}
        <span class="unread-toggle-label">${e.unreadOnly?"Unread":"All"}</span>
      </button>
    </div>`;if(e.messages.length===0)return`${o}<div class="empty-state">${e.loadingList?"Loading messages…":"No messages found."}</div>`;const c=a.map(r=>{const i=r.latest,u=i.from?.emailAddress.name||i.from?.emailAddress.address||"Unknown",h=r.hasUnread?"unread":"",U=e.selectedId===i.id||r.conversationId&&e.selectedConversationId===r.conversationId?"selected":"",g=r.hasUnread?'<span class="unread-dot" aria-label="Unread"></span>':"",I=i.hasAttachments?'<span class="att-badge" title="Has attachments">📎</span>':"",T=r.count>1?`<span class="thread-count">(${r.count})</span>`:"",H=e.checkedIds.has(i.id)?"checked":"";return`
        <div class="message-row ${h} ${U}">
          <input type="checkbox" class="msg-check" data-check-id="${d(i.id)}" ${H} aria-label="Select ${d(i.subject||"message")}" />
          <button class="message-item" data-id="${d(i.id)}" data-conv-id="${d(r.conversationId??"")}" data-thread-count="${r.count}">
            <div class="message-item-top">
              <span class="sender-row">
                ${g}
                <span class="sender-name">${d(u)}</span>
                ${T}
              </span>
              <span class="message-time">${I}${d(L(i.receivedDateTime))}</span>
            </div>
            <div class="message-subject">${d(i.subject||"(No subject)")}</div>
            <div class="message-preview">${d(i.bodyPreview)}</div>
          </button>
        </div>`}).join(""),m=e.hasMore?`<div class="load-more-wrap"><button id="load-more-btn" class="btn load-more-btn" ${e.loadingMore?"disabled":""}>${e.loadingMore?"Loading older…":"Load older messages"}</button></div>`:"";return o+c+m}function ue(){if(e.loadingDetail)return'<div class="empty-state">Loading message…</div>';const a=e.selectedDetail;if(!a)return'<div class="empty-state">Select a message to read or reply.</div>';const t=a.from?.emailAddress.name?`${a.from.emailAddress.name} <${a.from.emailAddress.address}>`:a.from?.emailAddress.address??"",s=a.body.contentType.toLowerCase()==="html"?a.body.content:`<pre class="email-body-text">${d(a.body.content)}</pre>`,l=a.attachments&&a.attachments.length>0?`
      <div class="detail-attachments">
        ${a.attachments.map(i=>`
            <button type="button" class="attachment-chip attachment-download" data-download-att="${d(i.id)}" data-att-name="${d(i.name)}">
              <span>📎 ${d(i.name)}</span>
              <span class="att-size">${d(je(i.size))}</span>
            </button>`).join("")}
      </div>`:"",n=`
    <details class="quote-preview">
      <summary>Quoted message (${d(L(a.receivedDateTime))})</summary>
      <div class="quote-meta">On ${d(L(a.receivedDateTime))}, ${d(t)} wrote:</div>
      <div class="quote-body">${s}</div>
    </details>`,o=e.replyOpen?`
      <div class="reply-box">
        <strong>${e.replyAll?"Reply All":"Reply"}</strong>
        <textarea id="reply-text" class="textarea" placeholder="Write your reply… (Cmd/Ctrl+Enter to send)"></textarea>
        <div class="attachment-picker-row">
          <label class="btn btn-sm file-picker-label">
            Attach files
            <input id="reply-files" type="file" multiple hidden />
          </label>
          <div id="reply-att-pills">${B(e.replyAttachments,"reply")}</div>
        </div>
        ${n}
        <div class="detail-actions">
          <button id="send-reply-btn" class="btn btn-primary" ${e.sendingReply?"disabled":""}>
            ${e.sendingReply?"Sending…":"Send Reply"}
          </button>
          <button id="cancel-reply-btn" class="btn">Cancel</button>
        </div>
      </div>`:"",c=e.forwardOpen?`
      <div class="reply-box">
        <strong>Forward</strong>
        <input id="forward-to" type="text" class="field-input" placeholder="To (comma-separated emails)" autocomplete="off" />
        <textarea id="forward-comment" class="textarea" placeholder="Add a note (optional)… (Cmd/Ctrl+Enter to send)"></textarea>
        ${n}
        <div class="detail-actions">
          <button id="send-forward-btn" class="btn btn-primary" ${e.sendingForward?"disabled":""}>
            ${e.sendingForward?"Forwarding…":"Send Forward"}
          </button>
          <button id="cancel-forward-btn" class="btn">Cancel</button>
        </div>
      </div>`:"",m=e.selectedThread.length>1?e.selectedThread.slice(0,-1).map(i=>{const u=i.from?.emailAddress.name||i.from?.emailAddress.address||"Unknown",h=i.body.contentType.toLowerCase()==="html"?i.body.content:`<pre class="email-body-text">${d(i.body.content)}</pre>`;return`
              <details class="thread-msg">
                <summary>
                  <span><strong>${d(u)}</strong></span>
                  <span class="message-time">${d(L(i.receivedDateTime))}</span>
                </summary>
                <div class="email-body">${h}</div>
              </details>`}).join(""):"",r=e.folder==="drafts"?'<button id="edit-draft-btn" class="btn btn-primary">Edit Draft</button>':"";return`
    ${m}
    <div class="detail-card">
      <div class="detail-header">
        <div>
          <h2 class="detail-subject">${d(a.subject||"(No subject)")}</h2>
          <div class="detail-meta">
            <span><strong>From:</strong> ${d(t)}</span>
            <span><strong>Date:</strong> ${d(L(a.receivedDateTime))}</span>
          </div>
        </div>
        <div class="detail-actions">
          ${r}
          <button id="toggle-read-btn" class="btn">${a.isRead?"Mark unread":"Mark read"}</button>
          <button id="reply-btn" class="btn btn-primary">Reply</button>
          <button id="reply-all-btn" class="btn">Reply All</button>
          <button id="forward-btn" class="btn">Forward</button>
          <button id="archive-btn" class="btn">Archive</button>
          <button id="delete-btn" class="btn btn-danger">Delete</button>
        </div>
      </div>
      ${l}
      <div class="email-body">${s}</div>
      ${o}
      ${c}
    </div>`}function Pe(a){const t=a.trim().split(/\s+/).filter(Boolean);return t.length===0?"?":t.length===1?t[0].slice(0,2).toUpperCase():(t[0][0]+t[t.length-1][0]).toUpperCase()}function Fe(a){const t=a.user;if(!t)return"";const s=t.organization||t.companyName||(t.mail.includes("@")?t.mail.split("@")[1].split(".")[0].toUpperCase():""),l=t.photoDataUrl?`<img class="avatar-img" src="${d(t.photoDataUrl)}" alt="${d(t.displayName)}" />`:`<span class="avatar-initials">${d(Pe(t.displayName))}</span>`,n=[s?`<div><span class="user-meta-k">Organization:</span> ${d(s)}</div>`:"",t.jobTitle?`<div><span class="user-meta-k">Title:</span> ${d(t.jobTitle)}</div>`:"",t.department?`<div><span class="user-meta-k">Department:</span> ${d(t.department)}</div>`:"",t.officeLocation?`<div><span class="user-meta-k">Office:</span> ${d(t.officeLocation)}</div>`:"",`<div><span class="user-meta-k">Mode:</span> ${d(a.mode)}</div>`].filter(Boolean).join(""),o=e.userMenuOpen?`
      <div id="user-popover" class="user-popover" role="dialog" aria-label="User details">
        <div class="user-popover-head">
          <div class="avatar-circle avatar-lg">${l}</div>
          <div class="user-popover-identity">
            <div class="user-popover-name">${d(t.displayName)}</div>
            <div class="user-popover-mail">${d(t.mail)}</div>
          </div>
        </div>
        <div class="user-popover-meta">${n}</div>
      </div>`:"";return`
    <div class="user-menu-wrap">
      ${s?`<span class="org-label" title="Organization">${d(s)}</span>`:""}
      <button type="button" id="user-menu-btn" class="avatar-btn" title="${d(t.displayName)} (${d(t.mail)}) — click for details" aria-expanded="${e.userMenuOpen}" aria-label="User profile">
        <span class="avatar-circle">${l}</span>
      </button>
      ${o}
    </div>`}function p(){if(e.serverUnreachable){A.innerHTML=`
      <div class="empty-state">
        <h2>Local Server Unavailable</h2>
        <p>Cannot reach the Outlook plugin backend at <code>${d(E||window.location.origin)}</code>. Make sure the local service is running.</p>
        <button id="retry-server-btn" class="btn btn-primary">Retry connection</button>
      </div>`;return}const{auth:a}=e;if(!a){A.innerHTML='<div class="empty-state">Loading Outlook Mail…</div>';return}if(!a.authenticated){const g=a.entraConfigured?e.inTeams?'<button id="teams-login-btn" class="btn btn-primary">Sign in with Microsoft</button>':`<a class="btn btn-primary" href="${E}/api/auth/login">Sign in with Microsoft</a>`:"<p>Configure <code>AZURE_CLIENT_ID</code> and <code>AZURE_CLIENT_SECRET</code> in <code>.env</code>, or enable <code>DEV_MOCK_GRAPH=true</code>.</p>";A.innerHTML=`
      <div class="empty-state">
        <h2>Outlook Mail for Teams</h2>
        <p>Sign in with your Microsoft 365 account to view and manage your Outlook mailbox.</p>
        ${g}
      </div>`;return}const t=Fe(a),s=a.mode==="session"?'<button id="logout-btn" class="btn">Sign out</button>':"",l=a.entraConfigured?'<button id="switch-account-btn" class="btn" title="Sign in with another Microsoft account">Switch account</button>':"",n=e.error?`<div class="error-banner" role="alert"><span>${d(e.error)}</span><button id="dismiss-err" class="btn">Dismiss</button></div>`:"",o=de.map(g=>{const I=e.folders.find(H=>H.id===g.id),T=I&&I.unreadItemCount>0?` <span class="folder-badge">${I.unreadItemCount}</span>`:"";return`
      <button role="tab" aria-selected="${e.folder===g.id}"
        class="folder-tab ${e.folder===g.id?"active":""}"
        data-folder="${g.id}">${g.label}${T}</button>`}).join(""),c=e.composeOpen?`
      <div class="modal-backdrop" role="dialog" aria-modal="true" aria-label="${e.composeDraftId?"Edit Draft":"New Mail"}">
        <form id="compose-form" class="modal">
          <div class="modal-head">
            <h3>${e.composeDraftId?"Edit Draft":"New Mail"}</h3>
            <span id="draft-status" class="draft-status">${d(e.draftStatus)}</span>
          </div>
          <div class="autocomplete-wrap">
            <input id="compose-to" type="text" class="field-input" placeholder="To (comma-separated emails)" autocomplete="off" value="${d(e.composeInitial.to)}" required />
            <div id="compose-suggestions" class="autocomplete-list" hidden></div>
          </div>
          <input id="compose-subject" type="text" class="field-input" placeholder="Subject" value="${d(e.composeInitial.subject)}" required />
          <textarea id="compose-body" class="textarea" placeholder="Message body… (Cmd/Ctrl+Enter to send)">${d(e.composeInitial.body)}</textarea>
          <div class="attachment-picker-row">
            <label class="btn btn-sm file-picker-label">
              Attach files
              <input id="compose-files" type="file" multiple hidden />
            </label>
            <div id="compose-att-pills">${B(e.composeAttachments,"compose")}</div>
          </div>
          <div class="modal-footer">
            <button type="button" id="cancel-compose-btn" class="btn">Cancel</button>
            <button type="button" id="save-draft-btn" class="btn">Save Draft</button>
            <button type="submit" class="btn btn-primary" ${e.sendingCompose?"disabled":""}>
              ${e.sendingCompose?"Sending…":"Send"}
            </button>
          </div>
        </form>
      </div>`:"",m=$>0,r=$>=0&&$<k.length-1,i=e.versionInfo?`v${d(e.versionInfo.version)}`:"v1.0.4",u=e.versionInfo?`v${d(e.versionInfo.version)} (${d(e.versionInfo.commit)})`:"v1.0.4",h=e.versionInfo?.updateAvailable?`${e.versionInfo.behindCount} update(s) available`:"Up to date",U=e.versionInfo?.updateAvailable?'<span class="update-dot" title="Update available"></span>':"";A.innerHTML=`
    <header class="topbar">
      <div class="topbar-left">
        <div class="nav-arrows">
          <button id="nav-back-btn" class="btn btn-sm" title="Back (Alt+←)" ${m?"":"disabled"}>←</button>
          <button id="nav-forward-btn" class="btn btn-sm" title="Forward (Alt+→)" ${r?"":"disabled"}>→</button>
        </div>
        <h1 class="app-title">Outlook Mail</h1>
        <span class="mode-badge">${e.inTeams?"Teams Tab":"Browser"} · ${a.mode==="mock"?"Mock Graph":"Graph API"}</span>
        <div class="app-brand" tabindex="0">
          <span class="mode-badge">${i} ${U}</span>
          <div class="version-popover">
            <div class="version-card">
              <div><strong>Outlook Mail</strong> ${u}</div>
              <div class="version-meta">${d(h)}</div>
              <button id="app-update-btn" class="btn btn-primary" ${e.updatingApp?"disabled":""}>
                ${e.updatingApp?"Updating…":"Update"}
              </button>
            </div>
          </div>
        </div>
      </div>
      <div class="topbar-search">
        <input id="search-input" type="search" class="search-input" placeholder="Search mail…"
          aria-label="Search mail" value="${d(e.search)}" />
      </div>
      <div class="topbar-actions">
        ${t}
        <button id="new-mail-btn" class="btn btn-primary">New Mail</button>
        <button id="refresh-btn" class="btn" ${e.loadingList?"disabled":""}>
          ${e.loadingList?"Refreshing…":"Refresh"}
        </button>
        <a class="btn" href="https://outlook.office.com/mail/" target="_blank" rel="noopener noreferrer">Open OWA ↗</a>
        ${l}
        ${s}
      </div>
    </header>
    ${n}
    <div class="toolbar">
      <div class="folder-tabs" role="tablist" aria-label="Mail folders">${o}</div>
    </div>
    <main class="workspace">
      <section id="message-list" class="message-list" aria-label="Messages">${R()}</section>
      <section id="detail-pane" class="detail-pane" aria-label="Message detail">${ue()}</section>
    </main>
    ${c}`}function f(){const a=document.getElementById("message-list"),t=document.getElementById("detail-pane");if(!a||!t){p();return}a.innerHTML=R(),t.innerHTML=ue();const s=document.getElementById("refresh-btn");s&&(s.disabled=e.loadingList,s.textContent=e.loadingList?"Refreshing…":"Refresh")}function _e(){for(const a of de){const t=document.querySelector(`.folder-tab[data-folder="${a.id}"]`);if(!t)continue;const s=e.folders.find(n=>n.id===a.id),l=s&&s.unreadItemCount>0?` <span class="folder-badge">${s.unreadItemCount}</span>`:"";t.innerHTML=`${a.label}${l}`}}async function me(){try{const a=await oe();e.folders=a.folders,_e()}catch{}}async function y(){V?.abort(),V=new AbortController,e.loadingList=!0,e.hasMore=!1,e.error=null,f();try{const[a]=await Promise.all([N({folder:e.folder,unreadOnly:e.unreadOnly,search:e.search},V.signal),me()]);e.messages=a.messages,e.hasMore=a.hasMore}catch(a){if(a.name==="AbortError")return;e.error=a instanceof Error?a.message:"Failed to load messages",p();return}finally{e.loadingList=!1}f()}async function pe(){if(e.loadingMore||e.loadingList||!e.hasMore)return;e.loadingMore=!0;const a=document.getElementById("message-list");a&&(a.innerHTML=R());try{const s=await N({folder:e.folder,unreadOnly:e.unreadOnly,search:e.search,skip:e.messages.length,top:25}),l=new Set(e.messages.map(o=>o.id)),n=s.messages.filter(o=>!l.has(o.id));e.messages=[...e.messages,...n],e.hasMore=s.hasMore&&n.length>0}catch(s){e.error=s instanceof Error?s.message:"Failed to load older messages"}finally{e.loadingMore=!1}const t=document.getElementById("message-list");t&&(t.innerHTML=R())}function fe(){const a=document.getElementById("nav-back-btn"),t=document.getElementById("nav-forward-btn");a&&(a.disabled=$<=0),t&&(t.disabled=$<0||$>=k.length-1)}function x(){const a={folder:e.folder,unreadOnly:e.unreadOnly,search:e.search,selectedId:e.selectedId,conversationId:e.selectedConversationId},t=k[$];t&&t.folder===a.folder&&t.unreadOnly===a.unreadOnly&&t.search===a.search&&t.selectedId===a.selectedId&&t.conversationId===a.conversationId||(k.splice($+1),k.push(a),k.length>50&&k.shift(),$=k.length-1,fe())}async function _(a){const t=$+a;if(t<0||t>=k.length)return;$=t;const s=k[$],l=s.folder!==e.folder||s.unreadOnly!==e.unreadOnly||s.search!==e.search;e.folder=s.folder,e.unreadOnly=s.unreadOnly,e.search=s.search,e.checkedIds.clear(),l&&(e.selectedId=s.selectedId,e.selectedConversationId=s.conversationId,e.selectedDetail=null,e.selectedThread=[],p(),await y()),s.selectedId?await ge(s.selectedId,s.conversationId??void 0,!1):(e.selectedId=null,e.selectedConversationId=null,e.selectedDetail=null,e.selectedThread=[],f()),fe()}async function he(){const a=document.getElementById("compose-to")?.value??"",t=document.getElementById("compose-subject")?.value??"",s=document.getElementById("compose-body")?.value??"",l=a.split(",").map(o=>o.trim()).filter(Boolean);if(l.length===0&&!t.trim()&&!s.trim())return;e.draftStatus="Saving…";const n=document.getElementById("draft-status");n&&(n.textContent="Saving…");try{const o=await re({id:e.composeDraftId??void 0,to:l,subject:t,body:s});e.composeDraftId=o.id,e.draftStatus="Draft saved";const c=document.getElementById("draft-status");c&&(c.textContent="Draft saved")}catch{e.draftStatus="Draft save failed";const o=document.getElementById("draft-status");o&&(o.textContent="Draft save failed")}}async function ge(a,t,s=!0){j?.abort(),j=new AbortController,e.selectedId=a,e.selectedConversationId=t??null,e.selectedThread=[],e.replyOpen=!1,e.replyAttachments=[],e.forwardOpen=!1,e.error=null,s&&x();const l=O.get(a);if(l)e.selectedDetail=l,e.loadingDetail=!1,f();else{e.loadingDetail=!0,f();try{const o=await Ae(a,j.signal);O.set(a,o),e.selectedDetail=o}catch(o){if(o.name==="AbortError")return;e.error=o instanceof Error?o.message:"Failed to open message",e.loadingDetail=!1,p();return}finally{e.loadingDetail=!1}}const n=t||e.selectedDetail?.conversationId;n&&Te(n,j.signal).then(o=>{e.selectedId===a&&o.messages.length>1&&(e.selectedThread=o.messages,e.selectedDetail=o.messages[o.messages.length-1],f())}).catch(()=>{}),e.selectedDetail&&!e.selectedDetail.isRead?(e.selectedDetail.isRead=!0,e.messages=e.messages.map(o=>o.id===a?{...o,isRead:!0}:o),f(),ie(a,!0).catch(()=>{})):f()}A.addEventListener("click",async a=>{const t=a.target;if(t.closest("#user-menu-btn")){e.userMenuOpen=!e.userMenuOpen,p();return}if(e.userMenuOpen&&!t.closest(".user-menu-wrap")&&(e.userMenuOpen=!1,p()),t.closest(".msg-check"))return;if(t.id==="nav-back-btn"){_(-1);return}if(t.id==="nav-forward-btn"){_(1);return}const s=t.closest("[data-bulk]");if(s?.dataset.bulk&&e.checkedIds.size>0){const r=s.dataset.bulk,i=Array.from(e.checkedIds);try{await Oe(i,r),e.checkedIds.clear(),e.selectedId&&i.includes(e.selectedId)&&(r==="archive"||r==="delete")&&(e.selectedId=null,e.selectedDetail=null,e.selectedThread=[]),await y()}catch(u){e.error=u instanceof Error?u.message:"Bulk action failed",p()}return}if(t.closest("#bulk-clear-btn")){e.checkedIds.clear(),f();return}if(t.closest("#unread-filter-btn")){e.unreadOnly=!e.unreadOnly,x(),y();return}const n=t.closest(".message-item[data-id]");if(n?.dataset.id){const r=Number(n.dataset.threadCount??"1")>1&&n.dataset.convId||void 0;ge(n.dataset.id,r);return}const o=t.closest(".folder-tab[data-folder]");if(o?.dataset.folder){e.folder=o.dataset.folder,e.selectedId=null,e.selectedConversationId=null,e.selectedDetail=null,e.selectedThread=[],e.checkedIds.clear(),x(),p(),y();return}if(t.id==="retry-server-btn"){e.serverUnreachable=!1,e.auth=null,p(),Z();return}const c=t.closest("[data-download-att]");if(c?.dataset.downloadAtt&&e.selectedDetail){try{await De(e.selectedDetail.id,c.dataset.downloadAtt,c.dataset.attName??"attachment")}catch(r){e.error=r instanceof Error?r.message:"Failed to download attachment",p()}return}const m=t.closest("[data-remove-att]");if(m?.dataset.removeAtt){const r=Number(m.dataset.attIdx);if(m.dataset.removeAtt==="compose"){e.composeAttachments.splice(r,1);const i=document.getElementById("compose-att-pills");i&&(i.innerHTML=B(e.composeAttachments,"compose"))}else if(m.dataset.removeAtt==="reply"){e.replyAttachments.splice(r,1);const i=document.getElementById("reply-att-pills");i&&(i.innerHTML=B(e.replyAttachments,"reply"))}return}if(t.id==="load-more-btn"){pe();return}if(t.id==="refresh-btn")O.clear(),y();else if(t.id==="app-update-btn"){e.updatingApp=!0,p();try{await xe();for(let r=0;r<20;r++)if(await new Promise(u=>setTimeout(u,500)),await fetch(`${E}/api/health`).then(u=>u.ok).catch(()=>!1)){window.location.reload();return}window.location.reload()}catch(r){e.updatingApp=!1,e.error=r instanceof Error?r.message:"Update failed",p()}}else if(t.closest(".autocomplete-item[data-email]")){const r=t.closest(".autocomplete-item[data-email]");be(r.dataset.email)}else if(t.id==="new-mail-btn")e.composeOpen=!0,e.composeDraftId=null,e.composeInitial={to:"",subject:"",body:""},e.draftStatus="",e.composeAttachments=[],p();else if(t.id==="edit-draft-btn"&&e.selectedDetail){const r=e.selectedDetail,i=r.toRecipients.map(h=>h.emailAddress.address).filter(Boolean).join(", "),u=r.body.contentType.toLowerCase()==="html"?r.body.content.replace(/<[^>]+>/g," ").replace(/\s+/g," ").trim():r.body.content;e.composeOpen=!0,e.composeDraftId=r.id,e.composeInitial={to:i,subject:r.subject==="(Draft)"?"":r.subject,body:u},e.draftStatus="Loaded draft",e.composeAttachments=[],p()}else if(t.id==="save-draft-btn")he();else if(t.id==="cancel-compose-btn")M&&window.clearTimeout(M),e.composeOpen=!1,e.composeDraftId=null,e.composeAttachments=[],p();else if(t.id==="dismiss-err")e.error=null,p();else if(t.id==="teams-login-btn")try{const{authentication:r}=await F(async()=>{const{authentication:h}=await import("./src-CZ2ouLw7.js");return{authentication:h}},[],import.meta.url),i=await r.authenticate({url:`${E||window.location.origin}/api/auth/login?teams=1`,width:600,height:560});i&&q(i);const u=await P();e.auth=u,p(),u.authenticated&&y()}catch(r){e.error=r instanceof Error?r.message:"Microsoft sign-in failed",p()}else if(t.id==="switch-account-btn"){if(!e.inTeams){window.location.href=`${E}/api/auth/login?prompt=select_account`;return}try{const{authentication:r}=await F(async()=>{const{authentication:h}=await import("./src-CZ2ouLw7.js");return{authentication:h}},[],import.meta.url),i=await r.authenticate({url:`${E||window.location.origin}/api/auth/login?teams=1&prompt=select_account`,width:600,height:560});i&&q(i),O.clear(),e.selectedId=null,e.selectedConversationId=null,e.selectedDetail=null,e.selectedThread=[],e.checkedIds.clear();const u=await P();e.auth=u,p(),u.authenticated&&y()}catch(r){e.error=r instanceof Error?r.message:"Account switch failed",p()}}else if(t.id==="logout-btn")await ke(),q(null),O.clear(),e.auth=await P(),e.messages=[],e.selectedDetail=null,p();else if(t.id==="toggle-read-btn"&&e.selectedDetail){const r=e.selectedDetail,i=!r.isRead;r.isRead=i,e.messages=e.messages.map(u=>u.id===r.id?{...u,isRead:i}:u),f();try{await ie(r.id,i)}catch(u){e.error=u instanceof Error?u.message:"Failed to update read status",p()}}else if(t.id==="reply-btn")e.replyAll=!1,e.replyOpen=!0,e.forwardOpen=!1,f();else if(t.id==="reply-all-btn")e.replyAll=!0,e.replyOpen=!0,e.forwardOpen=!1,f();else if(t.id==="cancel-reply-btn")e.replyOpen=!1,e.replyAttachments=[],f();else if(t.id==="send-reply-btn"&&e.selectedDetail){const i=document.getElementById("reply-text")?.value.trim()??"";if(!i)return;e.sendingReply=!0,f();try{await Ce(e.selectedDetail.id,i,e.replyAll,e.replyAttachments.length>0?e.replyAttachments:void 0),e.replyOpen=!1,e.replyAttachments=[],e.sendingReply=!1,await y()}catch(u){e.sendingReply=!1,e.error=u instanceof Error?u.message:"Failed to send reply",p()}}else if(t.id==="forward-btn")e.forwardOpen=!0,e.replyOpen=!1,f();else if(t.id==="cancel-forward-btn")e.forwardOpen=!1,f();else if(t.id==="send-forward-btn"&&e.selectedDetail){const r=document.getElementById("forward-to")?.value??"",i=document.getElementById("forward-comment")?.value.trim()??"",u=r.split(",").map(h=>h.trim()).filter(Boolean);if(u.length===0){e.error="Enter at least one recipient email to forward",p();return}e.sendingForward=!0,f();try{await Se(e.selectedDetail.id,u,i),e.forwardOpen=!1,e.sendingForward=!1,await y()}catch(h){e.sendingForward=!1,e.error=h instanceof Error?h.message:"Failed to forward message",p()}}else if(t.id==="archive-btn"&&e.selectedDetail){const r=e.selectedDetail.id;try{await Me(r),O.delete(r),e.messages=e.messages.filter(i=>i.id!==r),e.selectedId=null,e.selectedDetail=null,f()}catch(i){e.error=i instanceof Error?i.message:"Failed to archive message",p()}}else if(t.id==="delete-btn"&&e.selectedDetail){const r=e.selectedDetail.id;try{await le(r),O.delete(r),e.messages=e.messages.filter(i=>i.id!==r),e.selectedId=null,e.selectedDetail=null,f()}catch(i){e.error=i instanceof Error?i.message:"Failed to delete message",p()}}});A.addEventListener("change",async a=>{const t=a.target;if(t.id==="select-all-chk"){const s=ce(e.messages);if(t.checked)for(const l of s)e.checkedIds.add(l.latest.id);else e.checkedIds.clear();f();return}if(t.classList.contains("msg-check")&&t.dataset.checkId){const s=t.dataset.checkId;t.checked?e.checkedIds.add(s):e.checkedIds.delete(s),f();return}if(t.id==="compose-files"){const s=await te(t.files);e.composeAttachments=[...e.composeAttachments,...s].slice(0,10);const l=document.getElementById("compose-att-pills");l&&(l.innerHTML=B(e.composeAttachments,"compose"))}else if(t.id==="reply-files"){const s=await te(t.files);e.replyAttachments=[...e.replyAttachments,...s].slice(0,10);const l=document.getElementById("reply-att-pills");l&&(l.innerHTML=B(e.replyAttachments,"reply"))}});function C(){const a=document.getElementById("compose-suggestions");if(a){if(v.length===0){a.hidden=!0,a.innerHTML="";return}a.hidden=!1,a.innerHTML=v.map((t,s)=>`
      <button type="button" class="autocomplete-item ${s===w?"active":""}" data-email="${d(t.address)}">
        <span><strong>${d(t.name)}</strong> &lt;${d(t.address)}&gt;</span>
        <span class="autocomplete-sub">${d(t.source)}</span>
      </button>`).join("")}}function be(a){const t=document.getElementById("compose-to");if(!t)return;const s=t.value.split(",");s[s.length-1]=` ${a}`,t.value=`${s.map(l=>l.trim()).filter(Boolean).join(", ")}, `,v=[],w=-1,C(),t.focus()}A.addEventListener("input",a=>{const t=a.target;if(t.id==="search-input")e.search=t.value,J&&window.clearTimeout(J),J=window.setTimeout(()=>{x(),y()},200);else if(t.id==="compose-to"||t.id==="compose-subject"||t.id==="compose-body"){if(M&&window.clearTimeout(M),M=window.setTimeout(()=>{e.composeOpen&&!e.sendingCompose&&he()},1800),t.id!=="compose-to")return;const s=t.value.split(",").pop()?.trim()??"";if(G&&window.clearTimeout(G),s.length<1){v=[],w=-1,C();return}G=window.setTimeout(async()=>{K?.abort(),K=new AbortController;try{v=(await Le(s,K.signal)).suggestions,w=v.length>0?0:-1,C()}catch{}},150)}});A.addEventListener("keydown",a=>{a.target.id!=="compose-to"||v.length===0||(a.key==="ArrowDown"?(a.preventDefault(),w=(w+1)%v.length,C()):a.key==="ArrowUp"?(a.preventDefault(),w=(w-1+v.length)%v.length,C()):a.key==="Enter"&&w>=0?(a.preventDefault(),be(v[w].address)):a.key==="Escape"&&(v=[],w=-1,C()))});A.addEventListener("submit",async a=>{const t=a.target;if(t.id!=="compose-form")return;a.preventDefault();const s=document.getElementById("compose-to")?.value??"",l=document.getElementById("compose-subject")?.value.trim()??"",n=document.getElementById("compose-body")?.value??"",o=s.split(",").map(m=>m.trim()).filter(Boolean);if(o.length===0||!l)return;e.sendingCompose=!0;const c=t.querySelector('button[type="submit"]');c&&(c.disabled=!0,c.textContent="Sending…");try{if(M&&window.clearTimeout(M),e.composeDraftId&&e.composeAttachments.length===0){const m=await re({id:e.composeDraftId,to:o,subject:l,body:n});await Ee(m.id)}else await Be({to:o,subject:l,body:n,attachments:e.composeAttachments.length>0?e.composeAttachments:void 0}),e.composeDraftId&&le(e.composeDraftId).catch(()=>{});e.composeOpen=!1,e.composeDraftId=null,e.composeAttachments=[],e.sendingCompose=!1,v=[],w=-1,p(),y()}catch(m){e.sendingCompose=!1,c&&(c.disabled=!1,c.textContent="Send"),e.error=m instanceof Error?m.message:"Failed to send email",p()}});async function ae(){if(!(!e.auth?.authenticated||document.hidden||e.composeOpen||e.replyOpen||e.loadingList||e.loadingMore))try{const[a]=await Promise.all([N({folder:e.folder,unreadOnly:e.unreadOnly,search:e.search}),me()]);if(e.messages.length<=25)e.messages=a.messages,e.hasMore=a.hasMore;else{const s=new Set(a.messages.map(n=>n.id)),l=e.messages.filter(n=>!s.has(n.id));e.messages=[...a.messages,...l]}const t=document.getElementById("message-list");if(t){const s=t.scrollTop;t.innerHTML=R(),t.scrollTop=s}}catch{}}const Ne=3e4;A.addEventListener("scroll",a=>{const t=a.target;t?.id!=="message-list"||!e.hasMore||e.loadingMore||e.loadingList||t.scrollTop+t.clientHeight>=t.scrollHeight-120&&pe()},!0);async function Z(){if(window.location.pathname==="/auth-done"){const t=new URLSearchParams(window.location.search).get("sid")??"";try{const{app:s,authentication:l}=await F(async()=>{const{app:n,authentication:o}=await import("./src-CZ2ouLw7.js");return{app:n,authentication:o}},[],import.meta.url);await s.initialize(),l.notifySuccess(t)}catch{window.close()}return}const a=await Ue(ee);e.inTeams=a.inTeams,ee(a.theme),$e(a.ssoToken);try{const t=P(),s=N({folder:e.folder,unreadOnly:e.unreadOnly,search:e.search}).catch(()=>null),l=oe().catch(()=>null),[n,o,c]=await Promise.all([t,s,l]);e.serverUnreachable=!1,e.auth=n,c&&(e.folders=c.folders),n.authenticated&&o?(e.messages=o.messages,e.hasMore=o.hasMore):n.authenticated&&!o&&y(),x(),p()}catch(t){if(t instanceof se){e.serverUnreachable=!0,p(),window.setTimeout(()=>{e.serverUnreachable&&Z()},4e3);return}e.error=t instanceof Error?t.message:"Failed to initialize",p()}Re(!0).then(t=>{e.versionInfo=t,p()}).catch(()=>{}),window.setInterval(()=>void ae(),Ne),document.addEventListener("visibilitychange",()=>{document.hidden||ae()}),document.addEventListener("keydown",t=>{t.altKey&&t.key==="ArrowLeft"||t.metaKey&&t.key==="["?(t.preventDefault(),_(-1)):t.altKey&&t.key==="ArrowRight"||t.metaKey&&t.key==="]"?(t.preventDefault(),_(1)):(t.metaKey||t.ctrlKey)&&t.key==="Enter"?e.composeOpen?(t.preventDefault(),document.getElementById("compose-form")?.requestSubmit()):e.replyOpen?(t.preventDefault(),document.getElementById("send-reply-btn")?.click()):e.forwardOpen&&(t.preventDefault(),document.getElementById("send-forward-btn")?.click()):t.key==="Escape"&&(e.composeOpen?(e.composeOpen=!1,e.composeAttachments=[],p()):(e.replyOpen||e.forwardOpen)&&(e.replyOpen=!1,e.replyAttachments=[],e.forwardOpen=!1,f()))})}Z();
