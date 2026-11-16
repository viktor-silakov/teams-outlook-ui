(function(){const a=document.createElement("link").relList;if(a&&a.supports&&a.supports("modulepreload"))return;for(const n of document.querySelectorAll('link[rel="modulepreload"]'))l(n);new MutationObserver(n=>{for(const o of n)if(o.type==="childList")for(const d of o.addedNodes)d.tagName==="LINK"&&d.rel==="modulepreload"&&l(d)}).observe(document,{childList:!0,subtree:!0});function s(n){const o={};return n.integrity&&(o.integrity=n.integrity),n.referrerPolicy&&(o.referrerPolicy=n.referrerPolicy),n.crossOrigin==="use-credentials"?o.credentials="include":n.crossOrigin==="anonymous"?o.credentials="omit":o.credentials="same-origin",o}function l(n){if(n.ep)return;n.ep=!0;const o=s(n);fetch(n.href,o)}})();const ye="modulepreload",ve=function(t,a){return new URL(t,a).href},Q={},G=function(a,s,l){let n=Promise.resolve();if(s&&s.length>0){const d=document.getElementsByTagName("link"),u=document.querySelector("meta[property=csp-nonce]"),r=u?.nonce||u?.getAttribute("nonce");n=Promise.allSettled(s.map(i=>{if(i=ve(i,l),i in Q)return;Q[i]=!0;const m=i.endsWith(".css"),g=m?'[rel="stylesheet"]':"";if(!!l)for(let E=d.length-1;E>=0;E--){const T=d[E];if(T.href===i&&(!m||T.rel==="stylesheet"))return}else if(document.querySelector(`link[href="${i}"]${g}`))return;const b=document.createElement("link");if(b.rel=m?"stylesheet":ye,m||(b.as="script"),b.crossOrigin="",b.href=i,r&&b.setAttribute("nonce",r),document.head.appendChild(b),m)return new Promise((E,T)=>{b.addEventListener("load",E),b.addEventListener("error",()=>T(new Error(`Unable to preload CSS for ${i}`)))})}))}function o(d){const u=new Event("vite:preloadError",{cancelable:!0});if(u.payload=d,window.dispatchEvent(u),!u.defaultPrevented)throw d}return n.then(d=>{for(const u of d||[])u.status==="rejected"&&o(u.reason);return a().catch(o)})},R=window.location.hostname==="localhost"||window.location.hostname==="127.0.0.1"?"":"https://localhost:3978",O="teams_outlook_sid";function we(){try{const t=localStorage.getItem(O);if(t)return t;const a=sessionStorage.getItem(O);if(a)return localStorage.setItem(O,a),a}catch{}return null}class ae extends Error{constructor(a="Local server at localhost:3978 is unreachable"){super(a),this.name="ServerUnreachableError"}}let z=null,K=we();function $e(t){z=t}function Z(t){K=t;try{t?(localStorage.setItem(O,t),sessionStorage.setItem(O,t)):(localStorage.removeItem(O),sessionStorage.removeItem(O))}catch{}}function Ie(t,a){const s=new Headers(t);return a&&!s.has("Content-Type")&&s.set("Content-Type","application/json"),z?(s.set("Authorization",`Bearer ${z}`),s.set("x-token-type","teams-sso")):K&&s.set("x-session-id",K),s}async function se(t,a={}){const s=Ie(a.headers,!!a.body);let l;try{l=await fetch(`${R}${t}`,{...a,headers:s,credentials:"include"})}catch(n){throw n instanceof DOMException&&n.name==="AbortError"?n:new ae}if(!l.ok){let n=`Request failed (${l.status})`;try{const o=await l.json();o.error?.message&&(n=o.error.message)}catch{}throw new Error(n)}return l}async function h(t,a={}){return await(await se(t,a)).json()}function J(){return h("/api/auth/status")}function ke(){return h("/api/auth/logout",{method:"POST"})}function ne(){return h("/api/mail/folders")}function F(t,a){const s=new URLSearchParams({folder:t.folder,unreadOnly:String(t.unreadOnly),top:String(t.top??25),skip:String(t.skip??0)});return t.search.trim()&&s.set("search",t.search.trim()),h(`/api/mail/messages?${s.toString()}`,{signal:a})}function Ae(t,a){return h(`/api/mail/messages/${encodeURIComponent(t)}`,{signal:a})}function Ee(t,a){return h(`/api/mail/conversations/${encodeURIComponent(t)}`,{signal:a})}function Te(t,a){return h("/api/mail/bulk",{method:"POST",body:JSON.stringify({ids:t,action:a})})}function oe(t){return h("/api/mail/drafts",{method:"POST",body:JSON.stringify(t)})}function Oe(t){return h(`/api/mail/drafts/${encodeURIComponent(t)}/send`,{method:"POST"})}function re(t,a){return h(`/api/mail/messages/${encodeURIComponent(t)}/read`,{method:"PATCH",body:JSON.stringify({isRead:a})})}async function De(t,a,s){const n=await(await se(`/api/mail/messages/${encodeURIComponent(t)}/attachments/${encodeURIComponent(a)}`)).blob(),o=URL.createObjectURL(n),d=document.createElement("a");d.href=o,d.download=s||"attachment",document.body.appendChild(d),d.click(),d.remove(),setTimeout(()=>URL.revokeObjectURL(o),1e3)}function ie(t){return h(`/api/mail/messages/${encodeURIComponent(t)}`,{method:"DELETE"})}function Se(t){return h(`/api/mail/messages/${encodeURIComponent(t)}/archive`,{method:"POST"})}function Me(t,a,s){return h(`/api/mail/messages/${encodeURIComponent(t)}/forward`,{method:"POST",body:JSON.stringify({to:a,comment:s})})}function Ce(t,a,s=!1,l){return h(`/api/mail/messages/${encodeURIComponent(t)}/reply`,{method:"POST",body:JSON.stringify({comment:a,replyAll:s,attachments:l})})}function Be(t){return h("/api/mail/messages/send",{method:"POST",body:JSON.stringify(t)})}function Le(t,a){return h(`/api/mail/recipients?q=${encodeURIComponent(t)}`,{signal:a})}function Re(t=!1){return h(`/api/app/version${t?"?check=1":""}`)}function xe(){return h("/api/app/update",{method:"POST"})}function Y(t){return t==="dark"||t==="contrast"?t:"default"}function H(t,a,s){return Promise.race([t,new Promise((l,n)=>setTimeout(()=>n(new Error(`${s} timeout`)),a))])}async function Ue(t){if(!(window.parent!==window||!!window.opener))return{inTeams:!1,theme:"default",ssoToken:null};try{const{app:s,authentication:l}=await G(async()=>{const{app:u,authentication:r}=await import("./src-CZ2ouLw7.js");return{app:u,authentication:r}},[],import.meta.url);await H(s.initialize(),500,"Teams init");const n=await H(s.getContext(),1e3,"Teams getContext"),o=Y(n.app?.theme);s.registerOnThemeChangeHandler(u=>{t(Y(u))});let d=null;try{d=await H(l.getAuthToken(),1200,"Teams getAuthToken")}catch{d=null}return{inTeams:!0,theme:o,ssoToken:d}}catch{return{inTeams:!1,theme:"default",ssoToken:null}}}const le=[{id:"inbox",label:"Inbox"},{id:"sentitems",label:"Sent"},{id:"drafts",label:"Drafts"},{id:"archive",label:"Archive"},{id:"deleteditems",label:"Deleted"}],A=document.getElementById("root"),M=new Map;let e={theme:"default",inTeams:!1,auth:null,serverUnreachable:!1,folder:"inbox",folders:[],unreadOnly:!1,search:"",messages:[],hasMore:!1,loadingMore:!1,selectedId:null,selectedConversationId:null,selectedDetail:null,selectedThread:[],checkedIds:new Set,loadingList:!1,loadingDetail:!1,replyOpen:!1,replyAll:!1,sendingReply:!1,replyAttachments:[],forwardOpen:!1,sendingForward:!1,composeOpen:!1,composeDraftId:null,composeInitial:{to:"",subject:"",body:""},draftStatus:"",sendingCompose:!1,composeAttachments:[],updatingApp:!1,versionInfo:null,error:null};const k=[];let w=-1,D=null,_=null,j=null,N=null,q=null,V=null,y=[],v=-1;function c(t){return t?t.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}function L(t){const a=new Date(t);return Number.isNaN(a.getTime())?t:a.toLocaleString([],{month:"short",day:"numeric",hour:"2-digit",minute:"2-digit"})}function X(t){e.theme=t,document.documentElement.setAttribute("data-theme",t)}function je(t){return t<1024?`${t} B`:t<1024*1024?`${(t/1024).toFixed(1)} KB`:`${(t/(1024*1024)).toFixed(1)} MB`}async function ee(t){if(!t||t.length===0)return[];const a=[];for(const s of Array.from(t).slice(0,10)){const l=await s.arrayBuffer(),n=new Uint8Array(l);let o="";const d=32768;for(let u=0;u<n.length;u+=d)o+=String.fromCharCode(...n.subarray(u,u+d));a.push({name:s.name,contentType:s.type||"application/octet-stream",contentBytes:btoa(o)})}return a}function B(t,a){return t.length===0?"":`
    <div class="attachment-chips">
      ${t.map((s,l)=>`
          <span class="attachment-chip">
            <span>${c(s.name)}</span>
            <button type="button" class="chip-remove" data-remove-att="${a}" data-att-idx="${l}" aria-label="Remove ${c(s.name)}">&times;</button>
          </span>`).join("")}
    </div>`}function de(t){const a=new Map;for(const s of t){const l=s.conversationId||s.id,n=a.get(l);n?(n.count+=1,s.isRead||(n.hasUnread=!0),n.ids.push(s.id)):a.set(l,{key:l,conversationId:s.conversationId,latest:s,count:1,hasUnread:!s.isRead,ids:[s.id]})}return Array.from(a.values())}const S={markRead:'<svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M2 6l6 4 6-4"/><path d="M2 6v6a1 1 0 001 1h10a1 1 0 001-1V6l-6-3.5L2 6z"/></svg>',markUnread:'<svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="3.5" width="12" height="9" rx="1"/><path d="M2.5 4l5.5 4 5.5-4"/></svg>',archive:'<svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="3" width="12" height="3" rx="0.5"/><path d="M3 6v6.5a1 1 0 001 1h8a1 1 0 001-1V6"/><path d="M6.5 9h3"/></svg>',delete:'<svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M2.5 4h11"/><path d="M6 4V2.5h4V4"/><path d="M4 4l.7 9a1 1 0 001 .9h4.6a1 1 0 001-.9L12 4"/></svg>',clear:'<svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M4 4l8 8M12 4l-8 8"/></svg>',unreadFilter:'<svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M2 4h12l-4.5 5v3.5l-3 1.5V9L2 4z"/></svg>'};function x(){const t=de(e.messages),a=e.checkedIds.size,s=a>0,l=t.length>0&&t.every(r=>e.checkedIds.has(r.latest.id)),n=s?"":"disabled",o=`
    <div class="bulk-bar" role="toolbar" aria-label="Message actions">
      <div class="bulk-left">
        <input id="select-all-chk" type="checkbox" class="msg-check-all" ${l?"checked":""} ${t.length===0?"disabled":""} title="${l?"Deselect all visible messages":"Select all visible messages"}" aria-label="Select all visible messages" />
        <span class="bulk-count">${s?`${a} selected`:""}</span>
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
    </div>`;if(e.messages.length===0)return`${o}<div class="empty-state">${e.loadingList?"Loading messages…":"No messages found."}</div>`;const d=t.map(r=>{const i=r.latest,m=i.from?.emailAddress.name||i.from?.emailAddress.address||"Unknown",g=r.hasUnread?"unread":"",I=e.selectedId===i.id||r.conversationId&&e.selectedConversationId===r.conversationId?"selected":"",b=r.hasUnread?'<span class="unread-dot" aria-label="Unread"></span>':"",E=i.hasAttachments?'<span class="att-badge" title="Has attachments">📎</span>':"",T=r.count>1?`<span class="thread-count">(${r.count})</span>`:"",be=e.checkedIds.has(i.id)?"checked":"";return`
        <div class="message-row ${g} ${I}">
          <input type="checkbox" class="msg-check" data-check-id="${c(i.id)}" ${be} aria-label="Select ${c(i.subject||"message")}" />
          <button class="message-item" data-id="${c(i.id)}" data-conv-id="${c(r.conversationId??"")}" data-thread-count="${r.count}">
            <div class="message-item-top">
              <span class="sender-row">
                ${b}
                <span class="sender-name">${c(m)}</span>
                ${T}
              </span>
              <span class="message-time">${E}${c(L(i.receivedDateTime))}</span>
            </div>
            <div class="message-subject">${c(i.subject||"(No subject)")}</div>
            <div class="message-preview">${c(i.bodyPreview)}</div>
          </button>
        </div>`}).join(""),u=e.hasMore?`<div class="load-more-wrap"><button id="load-more-btn" class="btn load-more-btn" ${e.loadingMore?"disabled":""}>${e.loadingMore?"Loading older…":"Load older messages"}</button></div>`:"";return o+d+u}function ce(){if(e.loadingDetail)return'<div class="empty-state">Loading message…</div>';const t=e.selectedDetail;if(!t)return'<div class="empty-state">Select a message to read or reply.</div>';const a=t.from?.emailAddress.name?`${t.from.emailAddress.name} <${t.from.emailAddress.address}>`:t.from?.emailAddress.address??"",s=t.body.contentType.toLowerCase()==="html"?t.body.content:`<pre class="email-body-text">${c(t.body.content)}</pre>`,l=t.attachments&&t.attachments.length>0?`
      <div class="detail-attachments">
        ${t.attachments.map(i=>`
            <button type="button" class="attachment-chip attachment-download" data-download-att="${c(i.id)}" data-att-name="${c(i.name)}">
              <span>📎 ${c(i.name)}</span>
              <span class="att-size">${c(je(i.size))}</span>
            </button>`).join("")}
      </div>`:"",n=`
    <details class="quote-preview">
      <summary>Quoted message (${c(L(t.receivedDateTime))})</summary>
      <div class="quote-meta">On ${c(L(t.receivedDateTime))}, ${c(a)} wrote:</div>
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
      </div>`:"",d=e.forwardOpen?`
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
      </div>`:"",u=e.selectedThread.length>1?e.selectedThread.slice(0,-1).map(i=>{const m=i.from?.emailAddress.name||i.from?.emailAddress.address||"Unknown",g=i.body.contentType.toLowerCase()==="html"?i.body.content:`<pre class="email-body-text">${c(i.body.content)}</pre>`;return`
              <details class="thread-msg">
                <summary>
                  <span><strong>${c(m)}</strong></span>
                  <span class="message-time">${c(L(i.receivedDateTime))}</span>
                </summary>
                <div class="email-body">${g}</div>
              </details>`}).join(""):"",r=e.folder==="drafts"?'<button id="edit-draft-btn" class="btn btn-primary">Edit Draft</button>':"";return`
    ${u}
    <div class="detail-card">
      <div class="detail-header">
        <div>
          <h2 class="detail-subject">${c(t.subject||"(No subject)")}</h2>
          <div class="detail-meta">
            <span><strong>From:</strong> ${c(a)}</span>
            <span><strong>Date:</strong> ${c(L(t.receivedDateTime))}</span>
          </div>
        </div>
        <div class="detail-actions">
          ${r}
          <button id="toggle-read-btn" class="btn">${t.isRead?"Mark unread":"Mark read"}</button>
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
      ${d}
    </div>`}function p(){if(e.serverUnreachable){A.innerHTML=`
      <div class="empty-state">
        <h2>Local Server Unavailable</h2>
        <p>Cannot reach the Outlook plugin backend at <code>${c(R||window.location.origin)}</code>. Make sure the local service is running.</p>
        <button id="retry-server-btn" class="btn btn-primary">Retry connection</button>
      </div>`;return}const{auth:t}=e;if(!t){A.innerHTML='<div class="empty-state">Loading Outlook Mail…</div>';return}if(!t.authenticated){const I=t.entraConfigured?e.inTeams?'<button id="teams-login-btn" class="btn btn-primary">Sign in with Microsoft</button>':`<a class="btn btn-primary" href="${R}/api/auth/login">Sign in with Microsoft</a>`:"<p>Configure <code>AZURE_CLIENT_ID</code> and <code>AZURE_CLIENT_SECRET</code> in <code>.env</code>, or enable <code>DEV_MOCK_GRAPH=true</code>.</p>";A.innerHTML=`
      <div class="empty-state">
        <h2>Outlook Mail for Teams</h2>
        <p>Sign in with your Microsoft 365 account to view and manage your Outlook mailbox.</p>
        ${I}
      </div>`;return}const a=t.user?`<span class="user-label">${c(t.user.displayName)} (${c(t.user.mail)})</span>`:"",s=t.mode==="session"?'<button id="logout-btn" class="btn">Sign out</button>':"",l=e.error?`<div class="error-banner" role="alert"><span>${c(e.error)}</span><button id="dismiss-err" class="btn">Dismiss</button></div>`:"",n=le.map(I=>{const b=e.folders.find(T=>T.id===I.id),E=b&&b.unreadItemCount>0?` <span class="folder-badge">${b.unreadItemCount}</span>`:"";return`
      <button role="tab" aria-selected="${e.folder===I.id}"
        class="folder-tab ${e.folder===I.id?"active":""}"
        data-folder="${I.id}">${I.label}${E}</button>`}).join(""),o=e.composeOpen?`
      <div class="modal-backdrop" role="dialog" aria-modal="true" aria-label="${e.composeDraftId?"Edit Draft":"New Mail"}">
        <form id="compose-form" class="modal">
          <div class="modal-head">
            <h3>${e.composeDraftId?"Edit Draft":"New Mail"}</h3>
            <span id="draft-status" class="draft-status">${c(e.draftStatus)}</span>
          </div>
          <div class="autocomplete-wrap">
            <input id="compose-to" type="text" class="field-input" placeholder="To (comma-separated emails)" autocomplete="off" value="${c(e.composeInitial.to)}" required />
            <div id="compose-suggestions" class="autocomplete-list" hidden></div>
          </div>
          <input id="compose-subject" type="text" class="field-input" placeholder="Subject" value="${c(e.composeInitial.subject)}" required />
          <textarea id="compose-body" class="textarea" placeholder="Message body… (Cmd/Ctrl+Enter to send)">${c(e.composeInitial.body)}</textarea>
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
      </div>`:"",d=w>0,u=w>=0&&w<k.length-1,r=e.versionInfo?`v${c(e.versionInfo.version)}`:"v1.0.4",i=e.versionInfo?`v${c(e.versionInfo.version)} (${c(e.versionInfo.commit)})`:"v1.0.4",m=e.versionInfo?.updateAvailable?`${e.versionInfo.behindCount} update(s) available`:"Up to date",g=e.versionInfo?.updateAvailable?'<span class="update-dot" title="Update available"></span>':"";A.innerHTML=`
    <header class="topbar">
      <div class="topbar-left">
        <div class="nav-arrows">
          <button id="nav-back-btn" class="btn btn-sm" title="Back (Alt+←)" ${d?"":"disabled"}>←</button>
          <button id="nav-forward-btn" class="btn btn-sm" title="Forward (Alt+→)" ${u?"":"disabled"}>→</button>
        </div>
        <h1 class="app-title">Outlook Mail</h1>
        <span class="mode-badge">${e.inTeams?"Teams Tab":"Browser"} · ${t.mode==="mock"?"Mock Graph":"Graph API"}</span>
        <div class="app-brand" tabindex="0">
          <span class="mode-badge">${r} ${g}</span>
          <div class="version-popover">
            <div class="version-card">
              <div><strong>Outlook Mail</strong> ${i}</div>
              <div class="version-meta">${c(m)}</div>
              <button id="app-update-btn" class="btn btn-primary" ${e.updatingApp?"disabled":""}>
                ${e.updatingApp?"Updating…":"Update"}
              </button>
            </div>
          </div>
        </div>
      </div>
      <div class="topbar-actions">
        ${a}
        <button id="new-mail-btn" class="btn btn-primary">New Mail</button>
        <button id="refresh-btn" class="btn" ${e.loadingList?"disabled":""}>
          ${e.loadingList?"Refreshing…":"Refresh"}
        </button>
        <a class="btn" href="https://outlook.office.com/mail/" target="_blank" rel="noopener noreferrer">Open OWA ↗</a>
        ${s}
      </div>
    </header>
    ${l}
    <div class="toolbar">
      <div class="folder-tabs" role="tablist" aria-label="Mail folders">${n}</div>
      <div class="search-controls">
        <input id="search-input" type="search" class="search-input" placeholder="Search mail…"
          aria-label="Search mail" value="${c(e.search)}" />
      </div>
    </div>
    <main class="workspace">
      <section id="message-list" class="message-list" aria-label="Messages">${x()}</section>
      <section id="detail-pane" class="detail-pane" aria-label="Message detail">${ce()}</section>
    </main>
    ${o}`}function f(){const t=document.getElementById("message-list"),a=document.getElementById("detail-pane");if(!t||!a){p();return}t.innerHTML=x(),a.innerHTML=ce();const s=document.getElementById("refresh-btn");s&&(s.disabled=e.loadingList,s.textContent=e.loadingList?"Refreshing…":"Refresh")}function Pe(){for(const t of le){const a=document.querySelector(`.folder-tab[data-folder="${t.id}"]`);if(!a)continue;const s=e.folders.find(n=>n.id===t.id),l=s&&s.unreadItemCount>0?` <span class="folder-badge">${s.unreadItemCount}</span>`:"";a.innerHTML=`${t.label}${l}`}}async function ue(){try{const t=await ne();e.folders=t.folders,Pe()}catch{}}async function $(){_?.abort(),_=new AbortController,e.loadingList=!0,e.hasMore=!1,e.error=null,f();try{const[t]=await Promise.all([F({folder:e.folder,unreadOnly:e.unreadOnly,search:e.search},_.signal),ue()]);e.messages=t.messages,e.hasMore=t.hasMore}catch(t){if(t.name==="AbortError")return;e.error=t instanceof Error?t.message:"Failed to load messages",p();return}finally{e.loadingList=!1}f()}async function me(){if(e.loadingMore||e.loadingList||!e.hasMore)return;e.loadingMore=!0;const t=document.getElementById("message-list");t&&(t.innerHTML=x());try{const s=await F({folder:e.folder,unreadOnly:e.unreadOnly,search:e.search,skip:e.messages.length,top:25}),l=new Set(e.messages.map(o=>o.id)),n=s.messages.filter(o=>!l.has(o.id));e.messages=[...e.messages,...n],e.hasMore=s.hasMore&&n.length>0}catch(s){e.error=s instanceof Error?s.message:"Failed to load older messages"}finally{e.loadingMore=!1}const a=document.getElementById("message-list");a&&(a.innerHTML=x())}function pe(){const t=document.getElementById("nav-back-btn"),a=document.getElementById("nav-forward-btn");t&&(t.disabled=w<=0),a&&(a.disabled=w<0||w>=k.length-1)}function U(){const t={folder:e.folder,unreadOnly:e.unreadOnly,search:e.search,selectedId:e.selectedId,conversationId:e.selectedConversationId},a=k[w];a&&a.folder===t.folder&&a.unreadOnly===t.unreadOnly&&a.search===t.search&&a.selectedId===t.selectedId&&a.conversationId===t.conversationId||(k.splice(w+1),k.push(t),k.length>50&&k.shift(),w=k.length-1,pe())}async function P(t){const a=w+t;if(a<0||a>=k.length)return;w=a;const s=k[w],l=s.folder!==e.folder||s.unreadOnly!==e.unreadOnly||s.search!==e.search;e.folder=s.folder,e.unreadOnly=s.unreadOnly,e.search=s.search,e.checkedIds.clear(),l&&(e.selectedId=s.selectedId,e.selectedConversationId=s.conversationId,e.selectedDetail=null,e.selectedThread=[],p(),await $()),s.selectedId?await he(s.selectedId,s.conversationId??void 0,!1):(e.selectedId=null,e.selectedConversationId=null,e.selectedDetail=null,e.selectedThread=[],f()),pe()}async function fe(){const t=document.getElementById("compose-to")?.value??"",a=document.getElementById("compose-subject")?.value??"",s=document.getElementById("compose-body")?.value??"",l=t.split(",").map(o=>o.trim()).filter(Boolean);if(l.length===0&&!a.trim()&&!s.trim())return;e.draftStatus="Saving…";const n=document.getElementById("draft-status");n&&(n.textContent="Saving…");try{const o=await oe({id:e.composeDraftId??void 0,to:l,subject:a,body:s});e.composeDraftId=o.id,e.draftStatus="Draft saved";const d=document.getElementById("draft-status");d&&(d.textContent="Draft saved")}catch{e.draftStatus="Draft save failed";const o=document.getElementById("draft-status");o&&(o.textContent="Draft save failed")}}async function he(t,a,s=!0){j?.abort(),j=new AbortController,e.selectedId=t,e.selectedConversationId=a??null,e.selectedThread=[],e.replyOpen=!1,e.replyAttachments=[],e.forwardOpen=!1,e.error=null,s&&U();const l=M.get(t);if(l)e.selectedDetail=l,e.loadingDetail=!1,f();else{e.loadingDetail=!0,f();try{const o=await Ae(t,j.signal);M.set(t,o),e.selectedDetail=o}catch(o){if(o.name==="AbortError")return;e.error=o instanceof Error?o.message:"Failed to open message",e.loadingDetail=!1,p();return}finally{e.loadingDetail=!1}}const n=a||e.selectedDetail?.conversationId;n&&Ee(n,j.signal).then(o=>{e.selectedId===t&&o.messages.length>1&&(e.selectedThread=o.messages,e.selectedDetail=o.messages[o.messages.length-1],f())}).catch(()=>{}),e.selectedDetail&&!e.selectedDetail.isRead?(e.selectedDetail.isRead=!0,e.messages=e.messages.map(o=>o.id===t?{...o,isRead:!0}:o),f(),re(t,!0).catch(()=>{})):f()}A.addEventListener("click",async t=>{const a=t.target;if(a.closest(".msg-check"))return;if(a.id==="nav-back-btn"){P(-1);return}if(a.id==="nav-forward-btn"){P(1);return}const s=a.closest("[data-bulk]");if(s?.dataset.bulk&&e.checkedIds.size>0){const r=s.dataset.bulk,i=Array.from(e.checkedIds);try{await Te(i,r),e.checkedIds.clear(),e.selectedId&&i.includes(e.selectedId)&&(r==="archive"||r==="delete")&&(e.selectedId=null,e.selectedDetail=null,e.selectedThread=[]),await $()}catch(m){e.error=m instanceof Error?m.message:"Bulk action failed",p()}return}if(a.closest("#bulk-clear-btn")){e.checkedIds.clear(),f();return}if(a.closest("#unread-filter-btn")){e.unreadOnly=!e.unreadOnly,U(),$();return}const n=a.closest(".message-item[data-id]");if(n?.dataset.id){const r=Number(n.dataset.threadCount??"1")>1&&n.dataset.convId||void 0;he(n.dataset.id,r);return}const o=a.closest(".folder-tab[data-folder]");if(o?.dataset.folder){e.folder=o.dataset.folder,e.selectedId=null,e.selectedConversationId=null,e.selectedDetail=null,e.selectedThread=[],e.checkedIds.clear(),U(),p(),$();return}if(a.id==="retry-server-btn"){e.serverUnreachable=!1,e.auth=null,p(),W();return}const d=a.closest("[data-download-att]");if(d?.dataset.downloadAtt&&e.selectedDetail){try{await De(e.selectedDetail.id,d.dataset.downloadAtt,d.dataset.attName??"attachment")}catch(r){e.error=r instanceof Error?r.message:"Failed to download attachment",p()}return}const u=a.closest("[data-remove-att]");if(u?.dataset.removeAtt){const r=Number(u.dataset.attIdx);if(u.dataset.removeAtt==="compose"){e.composeAttachments.splice(r,1);const i=document.getElementById("compose-att-pills");i&&(i.innerHTML=B(e.composeAttachments,"compose"))}else if(u.dataset.removeAtt==="reply"){e.replyAttachments.splice(r,1);const i=document.getElementById("reply-att-pills");i&&(i.innerHTML=B(e.replyAttachments,"reply"))}return}if(a.id==="load-more-btn"){me();return}if(a.id==="refresh-btn")M.clear(),$();else if(a.id==="app-update-btn"){e.updatingApp=!0,p();try{await xe();for(let r=0;r<20;r++)if(await new Promise(m=>setTimeout(m,500)),await fetch(`${R}/api/health`).then(m=>m.ok).catch(()=>!1)){window.location.reload();return}window.location.reload()}catch(r){e.updatingApp=!1,e.error=r instanceof Error?r.message:"Update failed",p()}}else if(a.closest(".autocomplete-item[data-email]")){const r=a.closest(".autocomplete-item[data-email]");ge(r.dataset.email)}else if(a.id==="new-mail-btn")e.composeOpen=!0,e.composeDraftId=null,e.composeInitial={to:"",subject:"",body:""},e.draftStatus="",e.composeAttachments=[],p();else if(a.id==="edit-draft-btn"&&e.selectedDetail){const r=e.selectedDetail,i=r.toRecipients.map(g=>g.emailAddress.address).filter(Boolean).join(", "),m=r.body.contentType.toLowerCase()==="html"?r.body.content.replace(/<[^>]+>/g," ").replace(/\s+/g," ").trim():r.body.content;e.composeOpen=!0,e.composeDraftId=r.id,e.composeInitial={to:i,subject:r.subject==="(Draft)"?"":r.subject,body:m},e.draftStatus="Loaded draft",e.composeAttachments=[],p()}else if(a.id==="save-draft-btn")fe();else if(a.id==="cancel-compose-btn")D&&window.clearTimeout(D),e.composeOpen=!1,e.composeDraftId=null,e.composeAttachments=[],p();else if(a.id==="dismiss-err")e.error=null,p();else if(a.id==="teams-login-btn")try{const{authentication:r}=await G(async()=>{const{authentication:g}=await import("./src-CZ2ouLw7.js");return{authentication:g}},[],import.meta.url),i=await r.authenticate({url:`${R||window.location.origin}/api/auth/login?teams=1`,width:600,height:560});i&&Z(i);const m=await J();e.auth=m,p(),m.authenticated&&$()}catch(r){e.error=r instanceof Error?r.message:"Microsoft sign-in failed",p()}else if(a.id==="logout-btn")await ke(),Z(null),M.clear(),e.auth=await J(),e.messages=[],e.selectedDetail=null,p();else if(a.id==="toggle-read-btn"&&e.selectedDetail){const r=e.selectedDetail,i=!r.isRead;r.isRead=i,e.messages=e.messages.map(m=>m.id===r.id?{...m,isRead:i}:m),f();try{await re(r.id,i)}catch(m){e.error=m instanceof Error?m.message:"Failed to update read status",p()}}else if(a.id==="reply-btn")e.replyAll=!1,e.replyOpen=!0,e.forwardOpen=!1,f();else if(a.id==="reply-all-btn")e.replyAll=!0,e.replyOpen=!0,e.forwardOpen=!1,f();else if(a.id==="cancel-reply-btn")e.replyOpen=!1,e.replyAttachments=[],f();else if(a.id==="send-reply-btn"&&e.selectedDetail){const i=document.getElementById("reply-text")?.value.trim()??"";if(!i)return;e.sendingReply=!0,f();try{await Ce(e.selectedDetail.id,i,e.replyAll,e.replyAttachments.length>0?e.replyAttachments:void 0),e.replyOpen=!1,e.replyAttachments=[],e.sendingReply=!1,await $()}catch(m){e.sendingReply=!1,e.error=m instanceof Error?m.message:"Failed to send reply",p()}}else if(a.id==="forward-btn")e.forwardOpen=!0,e.replyOpen=!1,f();else if(a.id==="cancel-forward-btn")e.forwardOpen=!1,f();else if(a.id==="send-forward-btn"&&e.selectedDetail){const r=document.getElementById("forward-to")?.value??"",i=document.getElementById("forward-comment")?.value.trim()??"",m=r.split(",").map(g=>g.trim()).filter(Boolean);if(m.length===0){e.error="Enter at least one recipient email to forward",p();return}e.sendingForward=!0,f();try{await Me(e.selectedDetail.id,m,i),e.forwardOpen=!1,e.sendingForward=!1,await $()}catch(g){e.sendingForward=!1,e.error=g instanceof Error?g.message:"Failed to forward message",p()}}else if(a.id==="archive-btn"&&e.selectedDetail){const r=e.selectedDetail.id;try{await Se(r),M.delete(r),e.messages=e.messages.filter(i=>i.id!==r),e.selectedId=null,e.selectedDetail=null,f()}catch(i){e.error=i instanceof Error?i.message:"Failed to archive message",p()}}else if(a.id==="delete-btn"&&e.selectedDetail){const r=e.selectedDetail.id;try{await ie(r),M.delete(r),e.messages=e.messages.filter(i=>i.id!==r),e.selectedId=null,e.selectedDetail=null,f()}catch(i){e.error=i instanceof Error?i.message:"Failed to delete message",p()}}});A.addEventListener("change",async t=>{const a=t.target;if(a.id==="select-all-chk"){const s=de(e.messages);if(a.checked)for(const l of s)e.checkedIds.add(l.latest.id);else e.checkedIds.clear();f();return}if(a.classList.contains("msg-check")&&a.dataset.checkId){const s=a.dataset.checkId;a.checked?e.checkedIds.add(s):e.checkedIds.delete(s),f();return}if(a.id==="compose-files"){const s=await ee(a.files);e.composeAttachments=[...e.composeAttachments,...s].slice(0,10);const l=document.getElementById("compose-att-pills");l&&(l.innerHTML=B(e.composeAttachments,"compose"))}else if(a.id==="reply-files"){const s=await ee(a.files);e.replyAttachments=[...e.replyAttachments,...s].slice(0,10);const l=document.getElementById("reply-att-pills");l&&(l.innerHTML=B(e.replyAttachments,"reply"))}});function C(){const t=document.getElementById("compose-suggestions");if(t){if(y.length===0){t.hidden=!0,t.innerHTML="";return}t.hidden=!1,t.innerHTML=y.map((a,s)=>`
      <button type="button" class="autocomplete-item ${s===v?"active":""}" data-email="${c(a.address)}">
        <span><strong>${c(a.name)}</strong> &lt;${c(a.address)}&gt;</span>
        <span class="autocomplete-sub">${c(a.source)}</span>
      </button>`).join("")}}function ge(t){const a=document.getElementById("compose-to");if(!a)return;const s=a.value.split(",");s[s.length-1]=` ${t}`,a.value=`${s.map(l=>l.trim()).filter(Boolean).join(", ")}, `,y=[],v=-1,C(),a.focus()}A.addEventListener("input",t=>{const a=t.target;if(a.id==="search-input")e.search=a.value,q&&window.clearTimeout(q),q=window.setTimeout(()=>{U(),$()},200);else if(a.id==="compose-to"||a.id==="compose-subject"||a.id==="compose-body"){if(D&&window.clearTimeout(D),D=window.setTimeout(()=>{e.composeOpen&&!e.sendingCompose&&fe()},1800),a.id!=="compose-to")return;const s=a.value.split(",").pop()?.trim()??"";if(V&&window.clearTimeout(V),s.length<1){y=[],v=-1,C();return}V=window.setTimeout(async()=>{N?.abort(),N=new AbortController;try{y=(await Le(s,N.signal)).suggestions,v=y.length>0?0:-1,C()}catch{}},150)}});A.addEventListener("keydown",t=>{t.target.id!=="compose-to"||y.length===0||(t.key==="ArrowDown"?(t.preventDefault(),v=(v+1)%y.length,C()):t.key==="ArrowUp"?(t.preventDefault(),v=(v-1+y.length)%y.length,C()):t.key==="Enter"&&v>=0?(t.preventDefault(),ge(y[v].address)):t.key==="Escape"&&(y=[],v=-1,C()))});A.addEventListener("submit",async t=>{const a=t.target;if(a.id!=="compose-form")return;t.preventDefault();const s=document.getElementById("compose-to")?.value??"",l=document.getElementById("compose-subject")?.value.trim()??"",n=document.getElementById("compose-body")?.value??"",o=s.split(",").map(u=>u.trim()).filter(Boolean);if(o.length===0||!l)return;e.sendingCompose=!0;const d=a.querySelector('button[type="submit"]');d&&(d.disabled=!0,d.textContent="Sending…");try{if(D&&window.clearTimeout(D),e.composeDraftId&&e.composeAttachments.length===0){const u=await oe({id:e.composeDraftId,to:o,subject:l,body:n});await Oe(u.id)}else await Be({to:o,subject:l,body:n,attachments:e.composeAttachments.length>0?e.composeAttachments:void 0}),e.composeDraftId&&ie(e.composeDraftId).catch(()=>{});e.composeOpen=!1,e.composeDraftId=null,e.composeAttachments=[],e.sendingCompose=!1,y=[],v=-1,p(),$()}catch(u){e.sendingCompose=!1,d&&(d.disabled=!1,d.textContent="Send"),e.error=u instanceof Error?u.message:"Failed to send email",p()}});async function te(){if(!(!e.auth?.authenticated||document.hidden||e.composeOpen||e.replyOpen||e.loadingList||e.loadingMore))try{const[t]=await Promise.all([F({folder:e.folder,unreadOnly:e.unreadOnly,search:e.search}),ue()]);if(e.messages.length<=25)e.messages=t.messages,e.hasMore=t.hasMore;else{const s=new Set(t.messages.map(n=>n.id)),l=e.messages.filter(n=>!s.has(n.id));e.messages=[...t.messages,...l]}const a=document.getElementById("message-list");if(a){const s=a.scrollTop;a.innerHTML=x(),a.scrollTop=s}}catch{}}const Fe=3e4;A.addEventListener("scroll",t=>{const a=t.target;a?.id!=="message-list"||!e.hasMore||e.loadingMore||e.loadingList||a.scrollTop+a.clientHeight>=a.scrollHeight-120&&me()},!0);async function W(){if(window.location.pathname==="/auth-done"){const a=new URLSearchParams(window.location.search).get("sid")??"";try{const{app:s,authentication:l}=await G(async()=>{const{app:n,authentication:o}=await import("./src-CZ2ouLw7.js");return{app:n,authentication:o}},[],import.meta.url);await s.initialize(),l.notifySuccess(a)}catch{window.close()}return}const t=await Ue(X);e.inTeams=t.inTeams,X(t.theme),$e(t.ssoToken);try{const a=J(),s=F({folder:e.folder,unreadOnly:e.unreadOnly,search:e.search}).catch(()=>null),l=ne().catch(()=>null),[n,o,d]=await Promise.all([a,s,l]);e.serverUnreachable=!1,e.auth=n,d&&(e.folders=d.folders),n.authenticated&&o?(e.messages=o.messages,e.hasMore=o.hasMore):n.authenticated&&!o&&$(),U(),p()}catch(a){if(a instanceof ae){e.serverUnreachable=!0,p(),window.setTimeout(()=>{e.serverUnreachable&&W()},4e3);return}e.error=a instanceof Error?a.message:"Failed to initialize",p()}Re(!0).then(a=>{e.versionInfo=a,p()}).catch(()=>{}),window.setInterval(()=>void te(),Fe),document.addEventListener("visibilitychange",()=>{document.hidden||te()}),document.addEventListener("keydown",a=>{a.altKey&&a.key==="ArrowLeft"||a.metaKey&&a.key==="["?(a.preventDefault(),P(-1)):a.altKey&&a.key==="ArrowRight"||a.metaKey&&a.key==="]"?(a.preventDefault(),P(1)):(a.metaKey||a.ctrlKey)&&a.key==="Enter"?e.composeOpen?(a.preventDefault(),document.getElementById("compose-form")?.requestSubmit()):e.replyOpen?(a.preventDefault(),document.getElementById("send-reply-btn")?.click()):e.forwardOpen&&(a.preventDefault(),document.getElementById("send-forward-btn")?.click()):a.key==="Escape"&&(e.composeOpen?(e.composeOpen=!1,e.composeAttachments=[],p()):(e.replyOpen||e.forwardOpen)&&(e.replyOpen=!1,e.replyAttachments=[],e.forwardOpen=!1,f()))})}W();
