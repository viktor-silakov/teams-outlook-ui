(function(){const t=document.createElement("link").relList;if(t&&t.supports&&t.supports("modulepreload"))return;for(const o of document.querySelectorAll('link[rel="modulepreload"]'))i(o);new MutationObserver(o=>{for(const n of o)if(n.type==="childList")for(const d of n.addedNodes)d.tagName==="LINK"&&d.rel==="modulepreload"&&i(d)}).observe(document,{childList:!0,subtree:!0});function a(o){const n={};return o.integrity&&(n.integrity=o.integrity),o.referrerPolicy&&(n.referrerPolicy=o.referrerPolicy),o.crossOrigin==="use-credentials"?n.credentials="include":o.crossOrigin==="anonymous"?n.credentials="omit":n.credentials="same-origin",n}function i(o){if(o.ep)return;o.ep=!0;const n=a(o);fetch(o.href,n)}})();const he="modulepreload",ge=function(s,t){return new URL(s,t).href},W={},J=function(t,a,i){let o=Promise.resolve();if(a&&a.length>0){const d=document.getElementsByTagName("link"),r=document.querySelector("meta[property=csp-nonce]"),c=r?.nonce||r?.getAttribute("nonce");o=Promise.allSettled(a.map(l=>{if(l=ge(l,i),l in W)return;W[l]=!0;const p=l.endsWith(".css"),A=p?'[rel="stylesheet"]':"";if(!!i)for(let E=d.length-1;E>=0;E--){const D=d[E];if(D.href===l&&(!p||D.rel==="stylesheet"))return}else if(document.querySelector(`link[href="${l}"]${A}`))return;const b=document.createElement("link");if(b.rel=p?"stylesheet":he,p||(b.as="script"),b.crossOrigin="",b.href=l,c&&b.setAttribute("nonce",c),document.head.appendChild(b),p)return new Promise((E,D)=>{b.addEventListener("load",E),b.addEventListener("error",()=>D(new Error(`Unable to preload CSS for ${l}`)))})}))}function n(d){const r=new Event("vite:preloadError",{cancelable:!0});if(r.payload=d,window.dispatchEvent(r),!r.defaultPrevented)throw d}return o.then(d=>{for(const r of d||[])r.status==="rejected"&&n(r.reason);return t().catch(n)})},B=window.location.hostname==="localhost"||window.location.hostname==="127.0.0.1"?"":"https://localhost:3978",T="teams_outlook_sid";function be(){try{const s=localStorage.getItem(T);if(s)return s;const t=sessionStorage.getItem(T);if(t)return localStorage.setItem(T,t),t}catch{}return null}class te extends Error{constructor(t="Local server at localhost:3978 is unreachable"){super(t),this.name="ServerUnreachableError"}}let z=null,K=be();function ye(s){z=s}function Q(s){K=s;try{s?(localStorage.setItem(T,s),sessionStorage.setItem(T,s)):(localStorage.removeItem(T),sessionStorage.removeItem(T))}catch{}}function ve(s,t){const a=new Headers(s);return t&&!a.has("Content-Type")&&a.set("Content-Type","application/json"),z?(a.set("Authorization",`Bearer ${z}`),a.set("x-token-type","teams-sso")):K&&a.set("x-session-id",K),a}async function se(s,t={}){const a=ve(t.headers,!!t.body);let i;try{i=await fetch(`${B}${s}`,{...t,headers:a,credentials:"include"})}catch(o){throw o instanceof DOMException&&o.name==="AbortError"?o:new te}if(!i.ok){let o=`Request failed (${i.status})`;try{const n=await i.json();n.error?.message&&(o=n.error.message)}catch{}throw new Error(o)}return i}async function h(s,t={}){return await(await se(s,t)).json()}function V(){return h("/api/auth/status")}function we(){return h("/api/auth/logout",{method:"POST"})}function ae(){return h("/api/mail/folders")}function F(s,t){const a=new URLSearchParams({folder:s.folder,unreadOnly:String(s.unreadOnly),top:String(s.top??25),skip:String(s.skip??0)});return s.search.trim()&&a.set("search",s.search.trim()),h(`/api/mail/messages?${a.toString()}`,{signal:t})}function Ie(s,t){return h(`/api/mail/messages/${encodeURIComponent(s)}`,{signal:t})}function $e(s,t){return h(`/api/mail/conversations/${encodeURIComponent(s)}`,{signal:t})}function ke(s,t){return h("/api/mail/bulk",{method:"POST",body:JSON.stringify({ids:s,action:t})})}function ne(s){return h("/api/mail/drafts",{method:"POST",body:JSON.stringify(s)})}function Ae(s){return h(`/api/mail/drafts/${encodeURIComponent(s)}/send`,{method:"POST"})}function oe(s,t){return h(`/api/mail/messages/${encodeURIComponent(s)}/read`,{method:"PATCH",body:JSON.stringify({isRead:t})})}async function Ee(s,t,a){const o=await(await se(`/api/mail/messages/${encodeURIComponent(s)}/attachments/${encodeURIComponent(t)}`)).blob(),n=URL.createObjectURL(o),d=document.createElement("a");d.href=n,d.download=a||"attachment",document.body.appendChild(d),d.click(),d.remove(),setTimeout(()=>URL.revokeObjectURL(n),1e3)}function re(s){return h(`/api/mail/messages/${encodeURIComponent(s)}`,{method:"DELETE"})}function Te(s){return h(`/api/mail/messages/${encodeURIComponent(s)}/archive`,{method:"POST"})}function Oe(s,t,a){return h(`/api/mail/messages/${encodeURIComponent(s)}/forward`,{method:"POST",body:JSON.stringify({to:t,comment:a})})}function De(s,t,a=!1,i){return h(`/api/mail/messages/${encodeURIComponent(s)}/reply`,{method:"POST",body:JSON.stringify({comment:t,replyAll:a,attachments:i})})}function Se(s){return h("/api/mail/messages/send",{method:"POST",body:JSON.stringify(s)})}function Me(s,t){return h(`/api/mail/recipients?q=${encodeURIComponent(s)}`,{signal:t})}function Ce(s=!1){return h(`/api/app/version${s?"?check=1":""}`)}function Le(){return h("/api/app/update",{method:"POST"})}function Z(s){return s==="dark"||s==="contrast"?s:"default"}function j(s,t,a){return Promise.race([s,new Promise((i,o)=>setTimeout(()=>o(new Error(`${a} timeout`)),t))])}async function Be(s){if(!(window.parent!==window||!!window.opener))return{inTeams:!1,theme:"default",ssoToken:null};try{const{app:a,authentication:i}=await J(async()=>{const{app:r,authentication:c}=await import("./src-CZ2ouLw7.js");return{app:r,authentication:c}},[],import.meta.url);await j(a.initialize(),500,"Teams init");const o=await j(a.getContext(),1e3,"Teams getContext"),n=Z(o.app?.theme);a.registerOnThemeChangeHandler(r=>{s(Z(r))});let d=null;try{d=await j(i.getAuthToken(),1200,"Teams getAuthToken")}catch{d=null}return{inTeams:!0,theme:n,ssoToken:d}}catch{return{inTeams:!1,theme:"default",ssoToken:null}}}const ie=[{id:"inbox",label:"Inbox"},{id:"sentitems",label:"Sent"},{id:"drafts",label:"Drafts"},{id:"archive",label:"Archive"},{id:"deleteditems",label:"Deleted"}],k=document.getElementById("root"),S=new Map;let e={theme:"default",inTeams:!1,auth:null,serverUnreachable:!1,folder:"inbox",folders:[],unreadOnly:!1,search:"",messages:[],hasMore:!1,loadingMore:!1,selectedId:null,selectedConversationId:null,selectedDetail:null,selectedThread:[],checkedIds:new Set,loadingList:!1,loadingDetail:!1,replyOpen:!1,replyAll:!1,sendingReply:!1,replyAttachments:[],forwardOpen:!1,sendingForward:!1,composeOpen:!1,composeDraftId:null,composeInitial:{to:"",subject:"",body:""},draftStatus:"",sendingCompose:!1,composeAttachments:[],updatingApp:!1,versionInfo:null,error:null};const $=[];let v=-1,O=null,H=null,U=null,_=null,N=null,q=null,g=[],y=-1;function u(s){return s?s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}function L(s){const t=new Date(s);return Number.isNaN(t.getTime())?s:t.toLocaleString([],{month:"short",day:"numeric",hour:"2-digit",minute:"2-digit"})}function Y(s){e.theme=s,document.documentElement.setAttribute("data-theme",s)}function Re(s){return s<1024?`${s} B`:s<1024*1024?`${(s/1024).toFixed(1)} KB`:`${(s/(1024*1024)).toFixed(1)} MB`}async function X(s){if(!s||s.length===0)return[];const t=[];for(const a of Array.from(s).slice(0,10)){const i=await a.arrayBuffer(),o=new Uint8Array(i);let n="";const d=32768;for(let r=0;r<o.length;r+=d)n+=String.fromCharCode(...o.subarray(r,r+d));t.push({name:a.name,contentType:a.type||"application/octet-stream",contentBytes:btoa(n)})}return t}function C(s,t){return s.length===0?"":`
    <div class="attachment-chips">
      ${s.map((a,i)=>`
          <span class="attachment-chip">
            <span>${u(a.name)}</span>
            <button type="button" class="chip-remove" data-remove-att="${t}" data-att-idx="${i}" aria-label="Remove ${u(a.name)}">&times;</button>
          </span>`).join("")}
    </div>`}function xe(s){const t=new Map;for(const a of s){const i=a.conversationId||a.id,o=t.get(i);o?(o.count+=1,a.isRead||(o.hasUnread=!0),o.ids.push(a.id)):t.set(i,{key:i,conversationId:a.conversationId,latest:a,count:1,hasUnread:!a.isRead,ids:[a.id]})}return Array.from(t.values())}function R(){if(e.messages.length===0)return`<div class="empty-state">${e.loadingList?"Loading messages…":"No messages found."}</div>`;const s=xe(e.messages),t=e.checkedIds.size>0?`
      <div class="bulk-bar">
        <span><strong>${e.checkedIds.size}</strong> selected</span>
        <div class="bulk-actions">
          <button type="button" class="btn btn-sm" data-bulk="markRead">Mark read</button>
          <button type="button" class="btn btn-sm" data-bulk="markUnread">Mark unread</button>
          <button type="button" class="btn btn-sm" data-bulk="archive">Archive</button>
          <button type="button" class="btn btn-sm btn-danger" data-bulk="delete">Delete</button>
          <button type="button" class="btn btn-sm" id="bulk-clear-btn">Clear</button>
        </div>
      </div>`:"",a=s.map(o=>{const n=o.latest,d=n.from?.emailAddress.name||n.from?.emailAddress.address||"Unknown",r=o.hasUnread?"unread":"",c=e.selectedId===n.id||o.conversationId&&e.selectedConversationId===o.conversationId?"selected":"",l=o.hasUnread?'<span class="unread-dot" aria-label="Unread"></span>':"",p=n.hasAttachments?'<span class="att-badge" title="Has attachments">📎</span>':"",A=o.count>1?`<span class="thread-count">(${o.count})</span>`:"",I=e.checkedIds.has(n.id)?"checked":"";return`
        <div class="message-row ${r} ${c}">
          <input type="checkbox" class="msg-check" data-check-id="${u(n.id)}" ${I} aria-label="Select ${u(n.subject||"message")}" />
          <button class="message-item" data-id="${u(n.id)}" data-conv-id="${u(o.conversationId??"")}" data-thread-count="${o.count}">
            <div class="message-item-top">
              <span class="sender-row">
                ${l}
                <span class="sender-name">${u(d)}</span>
                ${A}
              </span>
              <span class="message-time">${p}${u(L(n.receivedDateTime))}</span>
            </div>
            <div class="message-subject">${u(n.subject||"(No subject)")}</div>
            <div class="message-preview">${u(n.bodyPreview)}</div>
          </button>
        </div>`}).join(""),i=e.hasMore?`<div class="load-more-wrap"><button id="load-more-btn" class="btn load-more-btn" ${e.loadingMore?"disabled":""}>${e.loadingMore?"Loading older…":"Load older messages"}</button></div>`:"";return t+a+i}function le(){if(e.loadingDetail)return'<div class="empty-state">Loading message…</div>';const s=e.selectedDetail;if(!s)return'<div class="empty-state">Select a message to read or reply.</div>';const t=s.from?.emailAddress.name?`${s.from.emailAddress.name} <${s.from.emailAddress.address}>`:s.from?.emailAddress.address??"",a=s.body.contentType.toLowerCase()==="html"?s.body.content:`<pre class="email-body-text">${u(s.body.content)}</pre>`,i=s.attachments&&s.attachments.length>0?`
      <div class="detail-attachments">
        ${s.attachments.map(l=>`
            <button type="button" class="attachment-chip attachment-download" data-download-att="${u(l.id)}" data-att-name="${u(l.name)}">
              <span>📎 ${u(l.name)}</span>
              <span class="att-size">${u(Re(l.size))}</span>
            </button>`).join("")}
      </div>`:"",o=`
    <details class="quote-preview">
      <summary>Quoted message (${u(L(s.receivedDateTime))})</summary>
      <div class="quote-meta">On ${u(L(s.receivedDateTime))}, ${u(t)} wrote:</div>
      <div class="quote-body">${a}</div>
    </details>`,n=e.replyOpen?`
      <div class="reply-box">
        <strong>${e.replyAll?"Reply All":"Reply"}</strong>
        <textarea id="reply-text" class="textarea" placeholder="Write your reply… (Cmd/Ctrl+Enter to send)"></textarea>
        <div class="attachment-picker-row">
          <label class="btn btn-sm file-picker-label">
            Attach files
            <input id="reply-files" type="file" multiple hidden />
          </label>
          <div id="reply-att-pills">${C(e.replyAttachments,"reply")}</div>
        </div>
        ${o}
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
        ${o}
        <div class="detail-actions">
          <button id="send-forward-btn" class="btn btn-primary" ${e.sendingForward?"disabled":""}>
            ${e.sendingForward?"Forwarding…":"Send Forward"}
          </button>
          <button id="cancel-forward-btn" class="btn">Cancel</button>
        </div>
      </div>`:"",r=e.selectedThread.length>1?e.selectedThread.slice(0,-1).map(l=>{const p=l.from?.emailAddress.name||l.from?.emailAddress.address||"Unknown",A=l.body.contentType.toLowerCase()==="html"?l.body.content:`<pre class="email-body-text">${u(l.body.content)}</pre>`;return`
              <details class="thread-msg">
                <summary>
                  <span><strong>${u(p)}</strong></span>
                  <span class="message-time">${u(L(l.receivedDateTime))}</span>
                </summary>
                <div class="email-body">${A}</div>
              </details>`}).join(""):"",c=e.folder==="drafts"?'<button id="edit-draft-btn" class="btn btn-primary">Edit Draft</button>':"";return`
    ${r}
    <div class="detail-card">
      <div class="detail-header">
        <div>
          <h2 class="detail-subject">${u(s.subject||"(No subject)")}</h2>
          <div class="detail-meta">
            <span><strong>From:</strong> ${u(t)}</span>
            <span><strong>Date:</strong> ${u(L(s.receivedDateTime))}</span>
          </div>
        </div>
        <div class="detail-actions">
          ${c}
          <button id="toggle-read-btn" class="btn">${s.isRead?"Mark unread":"Mark read"}</button>
          <button id="reply-btn" class="btn btn-primary">Reply</button>
          <button id="reply-all-btn" class="btn">Reply All</button>
          <button id="forward-btn" class="btn">Forward</button>
          <button id="archive-btn" class="btn">Archive</button>
          <button id="delete-btn" class="btn btn-danger">Delete</button>
        </div>
      </div>
      ${i}
      <div class="email-body">${a}</div>
      ${n}
      ${d}
    </div>`}function m(){if(e.serverUnreachable){k.innerHTML=`
      <div class="empty-state">
        <h2>Local Server Unavailable</h2>
        <p>Cannot reach the Outlook plugin backend at <code>${u(B||window.location.origin)}</code>. Make sure the local service is running.</p>
        <button id="retry-server-btn" class="btn btn-primary">Retry connection</button>
      </div>`;return}const{auth:s}=e;if(!s){k.innerHTML='<div class="empty-state">Loading Outlook Mail…</div>';return}if(!s.authenticated){const I=s.entraConfigured?e.inTeams?'<button id="teams-login-btn" class="btn btn-primary">Sign in with Microsoft</button>':`<a class="btn btn-primary" href="${B}/api/auth/login">Sign in with Microsoft</a>`:"<p>Configure <code>AZURE_CLIENT_ID</code> and <code>AZURE_CLIENT_SECRET</code> in <code>.env</code>, or enable <code>DEV_MOCK_GRAPH=true</code>.</p>";k.innerHTML=`
      <div class="empty-state">
        <h2>Outlook Mail for Teams</h2>
        <p>Sign in with your Microsoft 365 account to view and manage your Outlook mailbox.</p>
        ${I}
      </div>`;return}const t=s.user?`<span class="user-label">${u(s.user.displayName)} (${u(s.user.mail)})</span>`:"",a=s.mode==="session"?'<button id="logout-btn" class="btn">Sign out</button>':"",i=e.error?`<div class="error-banner" role="alert"><span>${u(e.error)}</span><button id="dismiss-err" class="btn">Dismiss</button></div>`:"",o=ie.map(I=>{const b=e.folders.find(D=>D.id===I.id),E=b&&b.unreadItemCount>0?` <span class="folder-badge">${b.unreadItemCount}</span>`:"";return`
      <button role="tab" aria-selected="${e.folder===I.id}"
        class="folder-tab ${e.folder===I.id?"active":""}"
        data-folder="${I.id}">${I.label}${E}</button>`}).join(""),n=e.composeOpen?`
      <div class="modal-backdrop" role="dialog" aria-modal="true" aria-label="${e.composeDraftId?"Edit Draft":"New Mail"}">
        <form id="compose-form" class="modal">
          <div class="modal-head">
            <h3>${e.composeDraftId?"Edit Draft":"New Mail"}</h3>
            <span id="draft-status" class="draft-status">${u(e.draftStatus)}</span>
          </div>
          <div class="autocomplete-wrap">
            <input id="compose-to" type="text" class="field-input" placeholder="To (comma-separated emails)" autocomplete="off" value="${u(e.composeInitial.to)}" required />
            <div id="compose-suggestions" class="autocomplete-list" hidden></div>
          </div>
          <input id="compose-subject" type="text" class="field-input" placeholder="Subject" value="${u(e.composeInitial.subject)}" required />
          <textarea id="compose-body" class="textarea" placeholder="Message body… (Cmd/Ctrl+Enter to send)">${u(e.composeInitial.body)}</textarea>
          <div class="attachment-picker-row">
            <label class="btn btn-sm file-picker-label">
              Attach files
              <input id="compose-files" type="file" multiple hidden />
            </label>
            <div id="compose-att-pills">${C(e.composeAttachments,"compose")}</div>
          </div>
          <div class="modal-footer">
            <button type="button" id="cancel-compose-btn" class="btn">Cancel</button>
            <button type="button" id="save-draft-btn" class="btn">Save Draft</button>
            <button type="submit" class="btn btn-primary" ${e.sendingCompose?"disabled":""}>
              ${e.sendingCompose?"Sending…":"Send"}
            </button>
          </div>
        </form>
      </div>`:"",d=v>0,r=v>=0&&v<$.length-1,c=e.versionInfo?`v${u(e.versionInfo.version)}`:"v1.0.4",l=e.versionInfo?`v${u(e.versionInfo.version)} (${u(e.versionInfo.commit)})`:"v1.0.4",p=e.versionInfo?.updateAvailable?`${e.versionInfo.behindCount} update(s) available`:"Up to date",A=e.versionInfo?.updateAvailable?'<span class="update-dot" title="Update available"></span>':"";k.innerHTML=`
    <header class="topbar">
      <div class="topbar-left">
        <div class="nav-arrows">
          <button id="nav-back-btn" class="btn btn-sm" title="Back (Alt+←)" ${d?"":"disabled"}>←</button>
          <button id="nav-forward-btn" class="btn btn-sm" title="Forward (Alt+→)" ${r?"":"disabled"}>→</button>
        </div>
        <h1 class="app-title">Outlook Mail</h1>
        <span class="mode-badge">${e.inTeams?"Teams Tab":"Browser"} · ${s.mode==="mock"?"Mock Graph":"Graph API"}</span>
        <div class="app-brand" tabindex="0">
          <span class="mode-badge">${c} ${A}</span>
          <div class="version-popover">
            <div class="version-card">
              <div><strong>Outlook Mail</strong> ${l}</div>
              <div class="version-meta">${u(p)}</div>
              <button id="app-update-btn" class="btn btn-primary" ${e.updatingApp?"disabled":""}>
                ${e.updatingApp?"Updating…":"Update"}
              </button>
            </div>
          </div>
        </div>
      </div>
      <div class="topbar-actions">
        ${t}
        <button id="new-mail-btn" class="btn btn-primary">New Mail</button>
        <button id="refresh-btn" class="btn" ${e.loadingList?"disabled":""}>
          ${e.loadingList?"Refreshing…":"Refresh"}
        </button>
        <a class="btn" href="https://outlook.office.com/mail/" target="_blank" rel="noopener noreferrer">Open OWA ↗</a>
        ${a}
      </div>
    </header>
    ${i}
    <div class="toolbar">
      <div class="folder-tabs" role="tablist" aria-label="Mail folders">${o}</div>
      <div class="search-controls">
        <label class="checkbox-label">
          <input id="unread-toggle" type="checkbox" ${e.unreadOnly?"checked":""} />
          Unread only
        </label>
        <input id="search-input" type="search" class="search-input" placeholder="Search mail…"
          aria-label="Search mail" value="${u(e.search)}" />
      </div>
    </div>
    <main class="workspace">
      <section id="message-list" class="message-list" aria-label="Messages">${R()}</section>
      <section id="detail-pane" class="detail-pane" aria-label="Message detail">${le()}</section>
    </main>
    ${n}`}function f(){const s=document.getElementById("message-list"),t=document.getElementById("detail-pane");if(!s||!t){m();return}s.innerHTML=R(),t.innerHTML=le();const a=document.getElementById("refresh-btn");a&&(a.disabled=e.loadingList,a.textContent=e.loadingList?"Refreshing…":"Refresh")}function Ue(){for(const s of ie){const t=document.querySelector(`.folder-tab[data-folder="${s.id}"]`);if(!t)continue;const a=e.folders.find(o=>o.id===s.id),i=a&&a.unreadItemCount>0?` <span class="folder-badge">${a.unreadItemCount}</span>`:"";t.innerHTML=`${s.label}${i}`}}async function de(){try{const s=await ae();e.folders=s.folders,Ue()}catch{}}async function w(){H?.abort(),H=new AbortController,e.loadingList=!0,e.hasMore=!1,e.error=null,f();try{const[s]=await Promise.all([F({folder:e.folder,unreadOnly:e.unreadOnly,search:e.search},H.signal),de()]);e.messages=s.messages,e.hasMore=s.hasMore}catch(s){if(s.name==="AbortError")return;e.error=s instanceof Error?s.message:"Failed to load messages",m();return}finally{e.loadingList=!1}f()}async function ce(){if(e.loadingMore||e.loadingList||!e.hasMore)return;e.loadingMore=!0;const s=document.getElementById("message-list");s&&(s.innerHTML=R());try{const a=await F({folder:e.folder,unreadOnly:e.unreadOnly,search:e.search,skip:e.messages.length,top:25}),i=new Set(e.messages.map(n=>n.id)),o=a.messages.filter(n=>!i.has(n.id));e.messages=[...e.messages,...o],e.hasMore=a.hasMore&&o.length>0}catch(a){e.error=a instanceof Error?a.message:"Failed to load older messages"}finally{e.loadingMore=!1}const t=document.getElementById("message-list");t&&(t.innerHTML=R())}function ue(){const s=document.getElementById("nav-back-btn"),t=document.getElementById("nav-forward-btn");s&&(s.disabled=v<=0),t&&(t.disabled=v<0||v>=$.length-1)}function x(){const s={folder:e.folder,unreadOnly:e.unreadOnly,search:e.search,selectedId:e.selectedId,conversationId:e.selectedConversationId},t=$[v];t&&t.folder===s.folder&&t.unreadOnly===s.unreadOnly&&t.search===s.search&&t.selectedId===s.selectedId&&t.conversationId===s.conversationId||($.splice(v+1),$.push(s),$.length>50&&$.shift(),v=$.length-1,ue())}async function P(s){const t=v+s;if(t<0||t>=$.length)return;v=t;const a=$[v],i=a.folder!==e.folder||a.unreadOnly!==e.unreadOnly||a.search!==e.search;e.folder=a.folder,e.unreadOnly=a.unreadOnly,e.search=a.search,e.checkedIds.clear(),i&&(e.selectedId=a.selectedId,e.selectedConversationId=a.conversationId,e.selectedDetail=null,e.selectedThread=[],m(),await w()),a.selectedId?await fe(a.selectedId,a.conversationId??void 0,!1):(e.selectedId=null,e.selectedConversationId=null,e.selectedDetail=null,e.selectedThread=[],f()),ue()}async function me(){const s=document.getElementById("compose-to")?.value??"",t=document.getElementById("compose-subject")?.value??"",a=document.getElementById("compose-body")?.value??"",i=s.split(",").map(n=>n.trim()).filter(Boolean);if(i.length===0&&!t.trim()&&!a.trim())return;e.draftStatus="Saving…";const o=document.getElementById("draft-status");o&&(o.textContent="Saving…");try{const n=await ne({id:e.composeDraftId??void 0,to:i,subject:t,body:a});e.composeDraftId=n.id,e.draftStatus="Draft saved";const d=document.getElementById("draft-status");d&&(d.textContent="Draft saved")}catch{e.draftStatus="Draft save failed";const n=document.getElementById("draft-status");n&&(n.textContent="Draft save failed")}}async function fe(s,t,a=!0){U?.abort(),U=new AbortController,e.selectedId=s,e.selectedConversationId=t??null,e.selectedThread=[],e.replyOpen=!1,e.replyAttachments=[],e.forwardOpen=!1,e.error=null,a&&x();const i=S.get(s);if(i)e.selectedDetail=i,e.loadingDetail=!1,f();else{e.loadingDetail=!0,f();try{const n=await Ie(s,U.signal);S.set(s,n),e.selectedDetail=n}catch(n){if(n.name==="AbortError")return;e.error=n instanceof Error?n.message:"Failed to open message",e.loadingDetail=!1,m();return}finally{e.loadingDetail=!1}}const o=t||e.selectedDetail?.conversationId;o&&$e(o,U.signal).then(n=>{e.selectedId===s&&n.messages.length>1&&(e.selectedThread=n.messages,e.selectedDetail=n.messages[n.messages.length-1],f())}).catch(()=>{}),e.selectedDetail&&!e.selectedDetail.isRead?(e.selectedDetail.isRead=!0,e.messages=e.messages.map(n=>n.id===s?{...n,isRead:!0}:n),f(),oe(s,!0).catch(()=>{})):f()}k.addEventListener("click",async s=>{const t=s.target;if(t.closest(".msg-check"))return;if(t.id==="nav-back-btn"){P(-1);return}if(t.id==="nav-forward-btn"){P(1);return}const a=t.closest("[data-bulk]");if(a?.dataset.bulk&&e.checkedIds.size>0){const r=a.dataset.bulk,c=Array.from(e.checkedIds);try{await ke(c,r),e.checkedIds.clear(),e.selectedId&&c.includes(e.selectedId)&&(r==="archive"||r==="delete")&&(e.selectedId=null,e.selectedDetail=null,e.selectedThread=[]),await w()}catch(l){e.error=l instanceof Error?l.message:"Bulk action failed",m()}return}if(t.id==="bulk-clear-btn"){e.checkedIds.clear(),f();return}const i=t.closest(".message-item[data-id]");if(i?.dataset.id){const r=Number(i.dataset.threadCount??"1")>1&&i.dataset.convId||void 0;fe(i.dataset.id,r);return}const o=t.closest(".folder-tab[data-folder]");if(o?.dataset.folder){e.folder=o.dataset.folder,e.selectedId=null,e.selectedConversationId=null,e.selectedDetail=null,e.selectedThread=[],e.checkedIds.clear(),x(),m(),w();return}if(t.id==="retry-server-btn"){e.serverUnreachable=!1,e.auth=null,m(),G();return}const n=t.closest("[data-download-att]");if(n?.dataset.downloadAtt&&e.selectedDetail){try{await Ee(e.selectedDetail.id,n.dataset.downloadAtt,n.dataset.attName??"attachment")}catch(r){e.error=r instanceof Error?r.message:"Failed to download attachment",m()}return}const d=t.closest("[data-remove-att]");if(d?.dataset.removeAtt){const r=Number(d.dataset.attIdx);if(d.dataset.removeAtt==="compose"){e.composeAttachments.splice(r,1);const c=document.getElementById("compose-att-pills");c&&(c.innerHTML=C(e.composeAttachments,"compose"))}else if(d.dataset.removeAtt==="reply"){e.replyAttachments.splice(r,1);const c=document.getElementById("reply-att-pills");c&&(c.innerHTML=C(e.replyAttachments,"reply"))}return}if(t.id==="load-more-btn"){ce();return}if(t.id==="refresh-btn")S.clear(),w();else if(t.id==="app-update-btn"){e.updatingApp=!0,m();try{await Le();for(let r=0;r<20;r++)if(await new Promise(l=>setTimeout(l,500)),await fetch(`${B}/api/health`).then(l=>l.ok).catch(()=>!1)){window.location.reload();return}window.location.reload()}catch(r){e.updatingApp=!1,e.error=r instanceof Error?r.message:"Update failed",m()}}else if(t.closest(".autocomplete-item[data-email]")){const r=t.closest(".autocomplete-item[data-email]");pe(r.dataset.email)}else if(t.id==="new-mail-btn")e.composeOpen=!0,e.composeDraftId=null,e.composeInitial={to:"",subject:"",body:""},e.draftStatus="",e.composeAttachments=[],m();else if(t.id==="edit-draft-btn"&&e.selectedDetail){const r=e.selectedDetail,c=r.toRecipients.map(p=>p.emailAddress.address).filter(Boolean).join(", "),l=r.body.contentType.toLowerCase()==="html"?r.body.content.replace(/<[^>]+>/g," ").replace(/\s+/g," ").trim():r.body.content;e.composeOpen=!0,e.composeDraftId=r.id,e.composeInitial={to:c,subject:r.subject==="(Draft)"?"":r.subject,body:l},e.draftStatus="Loaded draft",e.composeAttachments=[],m()}else if(t.id==="save-draft-btn")me();else if(t.id==="cancel-compose-btn")O&&window.clearTimeout(O),e.composeOpen=!1,e.composeDraftId=null,e.composeAttachments=[],m();else if(t.id==="dismiss-err")e.error=null,m();else if(t.id==="teams-login-btn")try{const{authentication:r}=await J(async()=>{const{authentication:p}=await import("./src-CZ2ouLw7.js");return{authentication:p}},[],import.meta.url),c=await r.authenticate({url:`${B||window.location.origin}/api/auth/login?teams=1`,width:600,height:560});c&&Q(c);const l=await V();e.auth=l,m(),l.authenticated&&w()}catch(r){e.error=r instanceof Error?r.message:"Microsoft sign-in failed",m()}else if(t.id==="logout-btn")await we(),Q(null),S.clear(),e.auth=await V(),e.messages=[],e.selectedDetail=null,m();else if(t.id==="toggle-read-btn"&&e.selectedDetail){const r=e.selectedDetail,c=!r.isRead;r.isRead=c,e.messages=e.messages.map(l=>l.id===r.id?{...l,isRead:c}:l),f();try{await oe(r.id,c)}catch(l){e.error=l instanceof Error?l.message:"Failed to update read status",m()}}else if(t.id==="reply-btn")e.replyAll=!1,e.replyOpen=!0,e.forwardOpen=!1,f();else if(t.id==="reply-all-btn")e.replyAll=!0,e.replyOpen=!0,e.forwardOpen=!1,f();else if(t.id==="cancel-reply-btn")e.replyOpen=!1,e.replyAttachments=[],f();else if(t.id==="send-reply-btn"&&e.selectedDetail){const c=document.getElementById("reply-text")?.value.trim()??"";if(!c)return;e.sendingReply=!0,f();try{await De(e.selectedDetail.id,c,e.replyAll,e.replyAttachments.length>0?e.replyAttachments:void 0),e.replyOpen=!1,e.replyAttachments=[],e.sendingReply=!1,await w()}catch(l){e.sendingReply=!1,e.error=l instanceof Error?l.message:"Failed to send reply",m()}}else if(t.id==="forward-btn")e.forwardOpen=!0,e.replyOpen=!1,f();else if(t.id==="cancel-forward-btn")e.forwardOpen=!1,f();else if(t.id==="send-forward-btn"&&e.selectedDetail){const r=document.getElementById("forward-to")?.value??"",c=document.getElementById("forward-comment")?.value.trim()??"",l=r.split(",").map(p=>p.trim()).filter(Boolean);if(l.length===0){e.error="Enter at least one recipient email to forward",m();return}e.sendingForward=!0,f();try{await Oe(e.selectedDetail.id,l,c),e.forwardOpen=!1,e.sendingForward=!1,await w()}catch(p){e.sendingForward=!1,e.error=p instanceof Error?p.message:"Failed to forward message",m()}}else if(t.id==="archive-btn"&&e.selectedDetail){const r=e.selectedDetail.id;try{await Te(r),S.delete(r),e.messages=e.messages.filter(c=>c.id!==r),e.selectedId=null,e.selectedDetail=null,f()}catch(c){e.error=c instanceof Error?c.message:"Failed to archive message",m()}}else if(t.id==="delete-btn"&&e.selectedDetail){const r=e.selectedDetail.id;try{await re(r),S.delete(r),e.messages=e.messages.filter(c=>c.id!==r),e.selectedId=null,e.selectedDetail=null,f()}catch(c){e.error=c instanceof Error?c.message:"Failed to delete message",m()}}});k.addEventListener("change",async s=>{const t=s.target;if(t.classList.contains("msg-check")&&t.dataset.checkId){const a=t.dataset.checkId;t.checked?e.checkedIds.add(a):e.checkedIds.delete(a),f();return}if(t.id==="unread-toggle")e.unreadOnly=t.checked,x(),w();else if(t.id==="compose-files"){const a=await X(t.files);e.composeAttachments=[...e.composeAttachments,...a].slice(0,10);const i=document.getElementById("compose-att-pills");i&&(i.innerHTML=C(e.composeAttachments,"compose"))}else if(t.id==="reply-files"){const a=await X(t.files);e.replyAttachments=[...e.replyAttachments,...a].slice(0,10);const i=document.getElementById("reply-att-pills");i&&(i.innerHTML=C(e.replyAttachments,"reply"))}});function M(){const s=document.getElementById("compose-suggestions");if(s){if(g.length===0){s.hidden=!0,s.innerHTML="";return}s.hidden=!1,s.innerHTML=g.map((t,a)=>`
      <button type="button" class="autocomplete-item ${a===y?"active":""}" data-email="${u(t.address)}">
        <span><strong>${u(t.name)}</strong> &lt;${u(t.address)}&gt;</span>
        <span class="autocomplete-sub">${u(t.source)}</span>
      </button>`).join("")}}function pe(s){const t=document.getElementById("compose-to");if(!t)return;const a=t.value.split(",");a[a.length-1]=` ${s}`,t.value=`${a.map(i=>i.trim()).filter(Boolean).join(", ")}, `,g=[],y=-1,M(),t.focus()}k.addEventListener("input",s=>{const t=s.target;if(t.id==="search-input")e.search=t.value,N&&window.clearTimeout(N),N=window.setTimeout(()=>{x(),w()},200);else if(t.id==="compose-to"||t.id==="compose-subject"||t.id==="compose-body"){if(O&&window.clearTimeout(O),O=window.setTimeout(()=>{e.composeOpen&&!e.sendingCompose&&me()},1800),t.id!=="compose-to")return;const a=t.value.split(",").pop()?.trim()??"";if(q&&window.clearTimeout(q),a.length<1){g=[],y=-1,M();return}q=window.setTimeout(async()=>{_?.abort(),_=new AbortController;try{g=(await Me(a,_.signal)).suggestions,y=g.length>0?0:-1,M()}catch{}},150)}});k.addEventListener("keydown",s=>{s.target.id!=="compose-to"||g.length===0||(s.key==="ArrowDown"?(s.preventDefault(),y=(y+1)%g.length,M()):s.key==="ArrowUp"?(s.preventDefault(),y=(y-1+g.length)%g.length,M()):s.key==="Enter"&&y>=0?(s.preventDefault(),pe(g[y].address)):s.key==="Escape"&&(g=[],y=-1,M()))});k.addEventListener("submit",async s=>{const t=s.target;if(t.id!=="compose-form")return;s.preventDefault();const a=document.getElementById("compose-to")?.value??"",i=document.getElementById("compose-subject")?.value.trim()??"",o=document.getElementById("compose-body")?.value??"",n=a.split(",").map(r=>r.trim()).filter(Boolean);if(n.length===0||!i)return;e.sendingCompose=!0;const d=t.querySelector('button[type="submit"]');d&&(d.disabled=!0,d.textContent="Sending…");try{if(O&&window.clearTimeout(O),e.composeDraftId&&e.composeAttachments.length===0){const r=await ne({id:e.composeDraftId,to:n,subject:i,body:o});await Ae(r.id)}else await Se({to:n,subject:i,body:o,attachments:e.composeAttachments.length>0?e.composeAttachments:void 0}),e.composeDraftId&&re(e.composeDraftId).catch(()=>{});e.composeOpen=!1,e.composeDraftId=null,e.composeAttachments=[],e.sendingCompose=!1,g=[],y=-1,m(),w()}catch(r){e.sendingCompose=!1,d&&(d.disabled=!1,d.textContent="Send"),e.error=r instanceof Error?r.message:"Failed to send email",m()}});async function ee(){if(!(!e.auth?.authenticated||document.hidden||e.composeOpen||e.replyOpen||e.loadingList||e.loadingMore))try{const[s]=await Promise.all([F({folder:e.folder,unreadOnly:e.unreadOnly,search:e.search}),de()]);if(e.messages.length<=25)e.messages=s.messages,e.hasMore=s.hasMore;else{const a=new Set(s.messages.map(o=>o.id)),i=e.messages.filter(o=>!a.has(o.id));e.messages=[...s.messages,...i]}const t=document.getElementById("message-list");if(t){const a=t.scrollTop;t.innerHTML=R(),t.scrollTop=a}}catch{}}const Pe=3e4;k.addEventListener("scroll",s=>{const t=s.target;t?.id!=="message-list"||!e.hasMore||e.loadingMore||e.loadingList||t.scrollTop+t.clientHeight>=t.scrollHeight-120&&ce()},!0);async function G(){if(window.location.pathname==="/auth-done"){const t=new URLSearchParams(window.location.search).get("sid")??"";try{const{app:a,authentication:i}=await J(async()=>{const{app:o,authentication:n}=await import("./src-CZ2ouLw7.js");return{app:o,authentication:n}},[],import.meta.url);await a.initialize(),i.notifySuccess(t)}catch{window.close()}return}const s=await Be(Y);e.inTeams=s.inTeams,Y(s.theme),ye(s.ssoToken);try{const t=V(),a=F({folder:e.folder,unreadOnly:e.unreadOnly,search:e.search}).catch(()=>null),i=ae().catch(()=>null),[o,n,d]=await Promise.all([t,a,i]);e.serverUnreachable=!1,e.auth=o,d&&(e.folders=d.folders),o.authenticated&&n?(e.messages=n.messages,e.hasMore=n.hasMore):o.authenticated&&!n&&w(),x(),m()}catch(t){if(t instanceof te){e.serverUnreachable=!0,m(),window.setTimeout(()=>{e.serverUnreachable&&G()},4e3);return}e.error=t instanceof Error?t.message:"Failed to initialize",m()}Ce(!0).then(t=>{e.versionInfo=t,m()}).catch(()=>{}),window.setInterval(()=>void ee(),Pe),document.addEventListener("visibilitychange",()=>{document.hidden||ee()}),document.addEventListener("keydown",t=>{t.altKey&&t.key==="ArrowLeft"||t.metaKey&&t.key==="["?(t.preventDefault(),P(-1)):t.altKey&&t.key==="ArrowRight"||t.metaKey&&t.key==="]"?(t.preventDefault(),P(1)):(t.metaKey||t.ctrlKey)&&t.key==="Enter"?e.composeOpen?(t.preventDefault(),document.getElementById("compose-form")?.requestSubmit()):e.replyOpen?(t.preventDefault(),document.getElementById("send-reply-btn")?.click()):e.forwardOpen&&(t.preventDefault(),document.getElementById("send-forward-btn")?.click()):t.key==="Escape"&&(e.composeOpen?(e.composeOpen=!1,e.composeAttachments=[],m()):(e.replyOpen||e.forwardOpen)&&(e.replyOpen=!1,e.replyAttachments=[],e.forwardOpen=!1,f()))})}G();
