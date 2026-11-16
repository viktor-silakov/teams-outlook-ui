(function(){const a=document.createElement("link").relList;if(a&&a.supports&&a.supports("modulepreload"))return;for(const r of document.querySelectorAll('link[rel="modulepreload"]'))i(r);new MutationObserver(r=>{for(const l of r)if(l.type==="childList")for(const s of l.addedNodes)s.tagName==="LINK"&&s.rel==="modulepreload"&&i(s)}).observe(document,{childList:!0,subtree:!0});function n(r){const l={};return r.integrity&&(l.integrity=r.integrity),r.referrerPolicy&&(l.referrerPolicy=r.referrerPolicy),r.crossOrigin==="use-credentials"?l.credentials="include":r.crossOrigin==="anonymous"?l.credentials="omit":l.credentials="same-origin",l}function i(r){if(r.ep)return;r.ep=!0;const l=n(r);fetch(r.href,l)}})();const ee="modulepreload",te=function(t,a){return new URL(t,a).href},j={},P=function(a,n,i){let r=Promise.resolve();if(n&&n.length>0){const s=document.getElementsByTagName("link"),o=document.querySelector("meta[property=csp-nonce]"),u=o?.nonce||o?.getAttribute("nonce");r=Promise.allSettled(n.map(m=>{if(m=te(m,i),m in j)return;j[m]=!0;const g=m.endsWith(".css"),X=g?'[rel="stylesheet"]':"";if(!!i)for(let O=s.length-1;O>=0;O--){const I=s[O];if(I.href===m&&(!g||I.rel==="stylesheet"))return}else if(document.querySelector(`link[href="${m}"]${X}`))return;const w=document.createElement("link");if(w.rel=g?"stylesheet":ee,g||(w.as="script"),w.crossOrigin="",w.href=m,u&&w.setAttribute("nonce",u),document.head.appendChild(w),g)return new Promise((O,I)=>{w.addEventListener("load",O),w.addEventListener("error",()=>I(new Error(`Unable to preload CSS for ${m}`)))})}))}function l(s){const o=new Event("vite:preloadError",{cancelable:!0});if(o.payload=s,window.dispatchEvent(o),!o.defaultPrevented)throw s}return r.then(s=>{for(const o of s||[])o.status==="rejected"&&l(o.reason);return a().catch(l)})},S=window.location.hostname==="localhost"||window.location.hostname==="127.0.0.1"?"":"https://localhost:3978",$="teams_outlook_sid";function ae(){try{const t=localStorage.getItem($);if(t)return t;const a=sessionStorage.getItem($);if(a)return localStorage.setItem($,a),a}catch{}return null}class G extends Error{constructor(a="Local server at localhost:3978 is unreachable"){super(a),this.name="ServerUnreachableError"}}let x=null,B=ae();function ne(t){x=t}function H(t){B=t;try{t?(localStorage.setItem($,t),sessionStorage.setItem($,t)):(localStorage.removeItem($),sessionStorage.removeItem($))}catch{}}function se(t,a){const n=new Headers(t);return a&&!n.has("Content-Type")&&n.set("Content-Type","application/json"),x?(n.set("Authorization",`Bearer ${x}`),n.set("x-token-type","teams-sso")):B&&n.set("x-session-id",B),n}async function J(t,a={}){const n=se(a.headers,!!a.body);let i;try{i=await fetch(`${S}${t}`,{...a,headers:n,credentials:"include"})}catch(r){throw r instanceof DOMException&&r.name==="AbortError"?r:new G}if(!i.ok){let r=`Request failed (${i.status})`;try{const l=await i.json();l.error?.message&&(r=l.error.message)}catch{}throw new Error(r)}return i}async function h(t,a={}){return await(await J(t,a)).json()}function U(){return h("/api/auth/status")}function oe(){return h("/api/auth/logout",{method:"POST"})}function _(t,a){const n=new URLSearchParams({folder:t.folder,unreadOnly:String(t.unreadOnly),top:String(t.top??25),skip:String(t.skip??0)});return t.search.trim()&&n.set("search",t.search.trim()),h(`/api/mail/messages?${n.toString()}`,{signal:a})}function re(t,a){return h(`/api/mail/messages/${encodeURIComponent(t)}`,{signal:a})}function W(t,a){return h(`/api/mail/messages/${encodeURIComponent(t)}/read`,{method:"PATCH",body:JSON.stringify({isRead:a})})}async function ie(t,a,n){const r=await(await J(`/api/mail/messages/${encodeURIComponent(t)}/attachments/${encodeURIComponent(a)}`)).blob(),l=URL.createObjectURL(r),s=document.createElement("a");s.href=l,s.download=n||"attachment",document.body.appendChild(s),s.click(),s.remove(),setTimeout(()=>URL.revokeObjectURL(l),1e3)}function le(t){return h(`/api/mail/messages/${encodeURIComponent(t)}`,{method:"DELETE"})}function de(t){return h(`/api/mail/messages/${encodeURIComponent(t)}/archive`,{method:"POST"})}function ce(t,a,n){return h(`/api/mail/messages/${encodeURIComponent(t)}/forward`,{method:"POST",body:JSON.stringify({to:a,comment:n})})}function ue(t,a,n=!1,i){return h(`/api/mail/messages/${encodeURIComponent(t)}/reply`,{method:"POST",body:JSON.stringify({comment:a,replyAll:n,attachments:i})})}function me(t){return h("/api/mail/messages/send",{method:"POST",body:JSON.stringify(t)})}function pe(t,a){return h(`/api/mail/recipients?q=${encodeURIComponent(t)}`,{signal:a})}function fe(t=!1){return h(`/api/app/version${t?"?check=1":""}`)}function he(){return h("/api/app/update",{method:"POST"})}function q(t){return t==="dark"||t==="contrast"?t:"default"}function L(t,a,n){return Promise.race([t,new Promise((i,r)=>setTimeout(()=>r(new Error(`${n} timeout`)),a))])}async function ge(t){if(!(window.parent!==window||!!window.opener))return{inTeams:!1,theme:"default",ssoToken:null};try{const{app:n,authentication:i}=await P(async()=>{const{app:o,authentication:u}=await import("./src-CZ2ouLw7.js");return{app:o,authentication:u}},[],import.meta.url);await L(n.initialize(),500,"Teams init");const r=await L(n.getContext(),1e3,"Teams getContext"),l=q(r.app?.theme);n.registerOnThemeChangeHandler(o=>{t(q(o))});let s=null;try{s=await L(i.getAuthToken(),1200,"Teams getAuthToken")}catch{s=null}return{inTeams:!0,theme:l,ssoToken:s}}catch{return{inTeams:!1,theme:"default",ssoToken:null}}}const be=[{id:"inbox",label:"Inbox"},{id:"sentitems",label:"Sent"},{id:"drafts",label:"Drafts"},{id:"archive",label:"Archive"}],v=document.getElementById("root"),A=new Map;let e={theme:"default",inTeams:!1,auth:null,serverUnreachable:!1,folder:"inbox",unreadOnly:!1,search:"",messages:[],selectedId:null,selectedDetail:null,loadingList:!1,loadingDetail:!1,replyOpen:!1,replyAll:!1,sendingReply:!1,replyAttachments:[],forwardOpen:!1,sendingForward:!1,composeOpen:!1,sendingCompose:!1,composeAttachments:[],updatingApp:!1,versionInfo:null,error:null},R=null,k=null,D=null,C=null,M=null,f=[],b=-1;function d(t){return t?t.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}function Z(t){const a=new Date(t);return Number.isNaN(a.getTime())?t:a.toLocaleString([],{month:"short",day:"numeric",hour:"2-digit",minute:"2-digit"})}function z(t){e.theme=t,document.documentElement.setAttribute("data-theme",t)}function ye(t){return t<1024?`${t} B`:t<1024*1024?`${(t/1024).toFixed(1)} KB`:`${(t/(1024*1024)).toFixed(1)} MB`}async function V(t){if(!t||t.length===0)return[];const a=[];for(const n of Array.from(t).slice(0,10)){const i=await n.arrayBuffer(),r=new Uint8Array(i);let l="";const s=32768;for(let o=0;o<r.length;o+=s)l+=String.fromCharCode(...r.subarray(o,o+s));a.push({name:n.name,contentType:n.type||"application/octet-stream",contentBytes:btoa(l)})}return a}function T(t,a){return t.length===0?"":`
    <div class="attachment-chips">
      ${t.map((n,i)=>`
          <span class="attachment-chip">
            <span>${d(n.name)}</span>
            <button type="button" class="chip-remove" data-remove-att="${a}" data-att-idx="${i}" aria-label="Remove ${d(n.name)}">&times;</button>
          </span>`).join("")}
    </div>`}function F(){return e.messages.length===0?`<div class="empty-state">${e.loadingList?"Loading messages…":"No messages found."}</div>`:e.messages.map(t=>{const a=t.from?.emailAddress.name||t.from?.emailAddress.address||"Unknown",n=t.isRead?"":"unread",i=e.selectedId===t.id?"selected":"",r=t.isRead?"":'<span class="unread-dot" aria-label="Unread"></span>',l=t.hasAttachments?'<span class="att-badge" title="Has attachments">📎</span>':"";return`
        <button class="message-item ${n} ${i}" data-id="${d(t.id)}">
          <div class="message-item-top">
            <span class="sender-row">
              ${r}
              <span class="sender-name">${d(a)}</span>
            </span>
            <span class="message-time">${l}${d(Z(t.receivedDateTime))}</span>
          </div>
          <div class="message-subject">${d(t.subject||"(No subject)")}</div>
          <div class="message-preview">${d(t.bodyPreview)}</div>
        </button>`}).join("")}function Q(){if(e.loadingDetail)return'<div class="empty-state">Loading message…</div>';const t=e.selectedDetail;if(!t)return'<div class="empty-state">Select a message to read or reply.</div>';const a=t.from?.emailAddress.name?`${t.from.emailAddress.name} <${t.from.emailAddress.address}>`:t.from?.emailAddress.address??"",n=t.body.contentType.toLowerCase()==="html"?t.body.content:`<pre class="email-body-text">${d(t.body.content)}</pre>`,i=t.attachments&&t.attachments.length>0?`
      <div class="detail-attachments">
        ${t.attachments.map(s=>`
            <button type="button" class="attachment-chip attachment-download" data-download-att="${d(s.id)}" data-att-name="${d(s.name)}">
              <span>📎 ${d(s.name)}</span>
              <span class="att-size">${d(ye(s.size))}</span>
            </button>`).join("")}
      </div>`:"",r=e.replyOpen?`
      <div class="reply-box">
        <strong>${e.replyAll?"Reply All":"Reply"}</strong>
        <textarea id="reply-text" class="textarea" placeholder="Write your reply… (Cmd/Ctrl+Enter to send)"></textarea>
        <div class="attachment-picker-row">
          <label class="btn btn-sm file-picker-label">
            Attach files
            <input id="reply-files" type="file" multiple hidden />
          </label>
          <div id="reply-att-pills">${T(e.replyAttachments,"reply")}</div>
        </div>
        <div class="detail-actions">
          <button id="send-reply-btn" class="btn btn-primary" ${e.sendingReply?"disabled":""}>
            ${e.sendingReply?"Sending…":"Send Reply"}
          </button>
          <button id="cancel-reply-btn" class="btn">Cancel</button>
        </div>
      </div>`:"",l=e.forwardOpen?`
      <div class="reply-box">
        <strong>Forward</strong>
        <input id="forward-to" type="text" class="field-input" placeholder="To (comma-separated emails)" autocomplete="off" />
        <textarea id="forward-comment" class="textarea" placeholder="Add a note (optional)… (Cmd/Ctrl+Enter to send)"></textarea>
        <div class="detail-actions">
          <button id="send-forward-btn" class="btn btn-primary" ${e.sendingForward?"disabled":""}>
            ${e.sendingForward?"Forwarding…":"Send Forward"}
          </button>
          <button id="cancel-forward-btn" class="btn">Cancel</button>
        </div>
      </div>`:"";return`
    <div class="detail-card">
      <div class="detail-header">
        <div>
          <h2 class="detail-subject">${d(t.subject||"(No subject)")}</h2>
          <div class="detail-meta">
            <span><strong>From:</strong> ${d(a)}</span>
            <span><strong>Date:</strong> ${d(Z(t.receivedDateTime))}</span>
          </div>
        </div>
        <div class="detail-actions">
          <button id="toggle-read-btn" class="btn">${t.isRead?"Mark unread":"Mark read"}</button>
          <button id="reply-btn" class="btn btn-primary">Reply</button>
          <button id="reply-all-btn" class="btn">Reply All</button>
          <button id="forward-btn" class="btn">Forward</button>
          <button id="archive-btn" class="btn">Archive</button>
          <button id="delete-btn" class="btn btn-danger">Delete</button>
        </div>
      </div>
      ${i}
      <div class="email-body">${n}</div>
      ${r}
      ${l}
    </div>`}function c(){if(e.serverUnreachable){v.innerHTML=`
      <div class="empty-state">
        <h2>Local Server Unavailable</h2>
        <p>Cannot reach the Outlook plugin backend at <code>${d(S||window.location.origin)}</code>. Make sure the local service is running.</p>
        <button id="retry-server-btn" class="btn btn-primary">Retry connection</button>
      </div>`;return}const{auth:t}=e;if(!t){v.innerHTML='<div class="empty-state">Loading Outlook Mail…</div>';return}if(!t.authenticated){const g=t.entraConfigured?e.inTeams?'<button id="teams-login-btn" class="btn btn-primary">Sign in with Microsoft</button>':`<a class="btn btn-primary" href="${S}/api/auth/login">Sign in with Microsoft</a>`:"<p>Configure <code>AZURE_CLIENT_ID</code> and <code>AZURE_CLIENT_SECRET</code> in <code>.env</code>, or enable <code>DEV_MOCK_GRAPH=true</code>.</p>";v.innerHTML=`
      <div class="empty-state">
        <h2>Outlook Mail for Teams</h2>
        <p>Sign in with your Microsoft 365 account to view and manage your Outlook mailbox.</p>
        ${g}
      </div>`;return}const a=t.user?`<span class="user-label">${d(t.user.displayName)} (${d(t.user.mail)})</span>`:"",n=t.mode==="session"?'<button id="logout-btn" class="btn">Sign out</button>':"",i=e.error?`<div class="error-banner" role="alert"><span>${d(e.error)}</span><button id="dismiss-err" class="btn">Dismiss</button></div>`:"",r=be.map(g=>`
      <button role="tab" aria-selected="${e.folder===g.id}"
        class="folder-tab ${e.folder===g.id?"active":""}"
        data-folder="${g.id}">${g.label}</button>`).join(""),l=e.composeOpen?`
      <div class="modal-backdrop" role="dialog" aria-modal="true" aria-label="New Mail">
        <form id="compose-form" class="modal">
          <h3>New Mail</h3>
          <div class="autocomplete-wrap">
            <input id="compose-to" type="text" class="field-input" placeholder="To (comma-separated emails)" autocomplete="off" required />
            <div id="compose-suggestions" class="autocomplete-list" hidden></div>
          </div>
          <input id="compose-subject" type="text" class="field-input" placeholder="Subject" required />
          <textarea id="compose-body" class="textarea" placeholder="Message body… (Cmd/Ctrl+Enter to send)"></textarea>
          <div class="attachment-picker-row">
            <label class="btn btn-sm file-picker-label">
              Attach files
              <input id="compose-files" type="file" multiple hidden />
            </label>
            <div id="compose-att-pills">${T(e.composeAttachments,"compose")}</div>
          </div>
          <div class="modal-footer">
            <button type="button" id="cancel-compose-btn" class="btn">Cancel</button>
            <button type="submit" class="btn btn-primary" ${e.sendingCompose?"disabled":""}>
              ${e.sendingCompose?"Sending…":"Send"}
            </button>
          </div>
        </form>
      </div>`:"",s=e.versionInfo?`v${d(e.versionInfo.version)}`:"v1.0.4",o=e.versionInfo?`v${d(e.versionInfo.version)} (${d(e.versionInfo.commit)})`:"v1.0.4",u=e.versionInfo?.updateAvailable?`${e.versionInfo.behindCount} update(s) available`:"Up to date",m=e.versionInfo?.updateAvailable?'<span class="update-dot" title="Update available"></span>':"";v.innerHTML=`
    <header class="topbar">
      <div class="topbar-left">
        <h1 class="app-title">Outlook Mail</h1>
        <span class="mode-badge">${e.inTeams?"Teams Tab":"Browser"} · ${t.mode==="mock"?"Mock Graph":"Graph API"}</span>
        <div class="app-brand" tabindex="0">
          <span class="mode-badge">${s} ${m}</span>
          <div class="version-popover">
            <div class="version-card">
              <div><strong>Outlook Mail</strong> ${o}</div>
              <div class="version-meta">${d(u)}</div>
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
        ${n}
      </div>
    </header>
    ${i}
    <div class="toolbar">
      <div class="folder-tabs" role="tablist" aria-label="Mail folders">${r}</div>
      <div class="search-controls">
        <label class="checkbox-label">
          <input id="unread-toggle" type="checkbox" ${e.unreadOnly?"checked":""} />
          Unread only
        </label>
        <input id="search-input" type="search" class="search-input" placeholder="Search mail…"
          aria-label="Search mail" value="${d(e.search)}" />
      </div>
    </div>
    <main class="workspace">
      <section id="message-list" class="message-list" aria-label="Messages">${F()}</section>
      <section id="detail-pane" class="detail-pane" aria-label="Message detail">${Q()}</section>
    </main>
    ${l}`}function p(){const t=document.getElementById("message-list"),a=document.getElementById("detail-pane");if(!t||!a){c();return}t.innerHTML=F(),a.innerHTML=Q();const n=document.getElementById("refresh-btn");n&&(n.disabled=e.loadingList,n.textContent=e.loadingList?"Refreshing…":"Refresh")}async function y(){R?.abort(),R=new AbortController,e.loadingList=!0,e.error=null,p();try{const t=await _({folder:e.folder,unreadOnly:e.unreadOnly,search:e.search},R.signal);e.messages=t.messages}catch(t){if(t.name==="AbortError")return;e.error=t instanceof Error?t.message:"Failed to load messages",c();return}finally{e.loadingList=!1}p()}async function ve(t){k?.abort(),k=new AbortController,e.selectedId=t,e.replyOpen=!1,e.replyAttachments=[],e.forwardOpen=!1,e.error=null;const a=A.get(t);if(a)e.selectedDetail=a,e.loadingDetail=!1,p();else{e.loadingDetail=!0,p();try{const n=await re(t,k.signal);A.set(t,n),e.selectedDetail=n}catch(n){if(n.name==="AbortError")return;e.error=n instanceof Error?n.message:"Failed to open message",e.loadingDetail=!1,c();return}finally{e.loadingDetail=!1}}e.selectedDetail&&!e.selectedDetail.isRead?(e.selectedDetail.isRead=!0,e.messages=e.messages.map(n=>n.id===t?{...n,isRead:!0}:n),p(),W(t,!0).catch(()=>{})):p()}v.addEventListener("click",async t=>{const a=t.target,n=a.closest(".message-item[data-id]");if(n?.dataset.id){ve(n.dataset.id);return}const i=a.closest(".folder-tab[data-folder]");if(i?.dataset.folder){e.folder=i.dataset.folder,e.selectedId=null,e.selectedDetail=null,c(),y();return}if(a.id==="retry-server-btn"){e.serverUnreachable=!1,e.auth=null,c(),N();return}const r=a.closest("[data-download-att]");if(r?.dataset.downloadAtt&&e.selectedDetail){try{await ie(e.selectedDetail.id,r.dataset.downloadAtt,r.dataset.attName??"attachment")}catch(s){e.error=s instanceof Error?s.message:"Failed to download attachment",c()}return}const l=a.closest("[data-remove-att]");if(l?.dataset.removeAtt){const s=Number(l.dataset.attIdx);if(l.dataset.removeAtt==="compose"){e.composeAttachments.splice(s,1);const o=document.getElementById("compose-att-pills");o&&(o.innerHTML=T(e.composeAttachments,"compose"))}else if(l.dataset.removeAtt==="reply"){e.replyAttachments.splice(s,1);const o=document.getElementById("reply-att-pills");o&&(o.innerHTML=T(e.replyAttachments,"reply"))}return}if(a.id==="refresh-btn")A.clear(),y();else if(a.id==="app-update-btn"){e.updatingApp=!0,c();try{await he();for(let s=0;s<20;s++)if(await new Promise(u=>setTimeout(u,500)),await fetch(`${S}/api/health`).then(u=>u.ok).catch(()=>!1)){window.location.reload();return}window.location.reload()}catch(s){e.updatingApp=!1,e.error=s instanceof Error?s.message:"Update failed",c()}}else if(a.closest(".autocomplete-item[data-email]")){const s=a.closest(".autocomplete-item[data-email]");Y(s.dataset.email)}else if(a.id==="new-mail-btn")e.composeOpen=!0,c();else if(a.id==="cancel-compose-btn")e.composeOpen=!1,c();else if(a.id==="dismiss-err")e.error=null,c();else if(a.id==="teams-login-btn")try{const{authentication:s}=await P(async()=>{const{authentication:m}=await import("./src-CZ2ouLw7.js");return{authentication:m}},[],import.meta.url),o=await s.authenticate({url:`${S||window.location.origin}/api/auth/login?teams=1`,width:600,height:560});o&&H(o);const u=await U();e.auth=u,c(),u.authenticated&&y()}catch(s){e.error=s instanceof Error?s.message:"Microsoft sign-in failed",c()}else if(a.id==="logout-btn")await oe(),H(null),A.clear(),e.auth=await U(),e.messages=[],e.selectedDetail=null,c();else if(a.id==="toggle-read-btn"&&e.selectedDetail){const s=e.selectedDetail,o=!s.isRead;s.isRead=o,e.messages=e.messages.map(u=>u.id===s.id?{...u,isRead:o}:u),p();try{await W(s.id,o)}catch(u){e.error=u instanceof Error?u.message:"Failed to update read status",c()}}else if(a.id==="reply-btn")e.replyAll=!1,e.replyOpen=!0,e.forwardOpen=!1,p();else if(a.id==="reply-all-btn")e.replyAll=!0,e.replyOpen=!0,e.forwardOpen=!1,p();else if(a.id==="cancel-reply-btn")e.replyOpen=!1,e.replyAttachments=[],p();else if(a.id==="send-reply-btn"&&e.selectedDetail){const o=document.getElementById("reply-text")?.value.trim()??"";if(!o)return;e.sendingReply=!0,p();try{await ue(e.selectedDetail.id,o,e.replyAll,e.replyAttachments.length>0?e.replyAttachments:void 0),e.replyOpen=!1,e.replyAttachments=[],e.sendingReply=!1,await y()}catch(u){e.sendingReply=!1,e.error=u instanceof Error?u.message:"Failed to send reply",c()}}else if(a.id==="forward-btn")e.forwardOpen=!0,e.replyOpen=!1,p();else if(a.id==="cancel-forward-btn")e.forwardOpen=!1,p();else if(a.id==="send-forward-btn"&&e.selectedDetail){const s=document.getElementById("forward-to")?.value??"",o=document.getElementById("forward-comment")?.value.trim()??"",u=s.split(",").map(m=>m.trim()).filter(Boolean);if(u.length===0){e.error="Enter at least one recipient email to forward",c();return}e.sendingForward=!0,p();try{await ce(e.selectedDetail.id,u,o),e.forwardOpen=!1,e.sendingForward=!1,await y()}catch(m){e.sendingForward=!1,e.error=m instanceof Error?m.message:"Failed to forward message",c()}}else if(a.id==="archive-btn"&&e.selectedDetail){const s=e.selectedDetail.id;try{await de(s),A.delete(s),e.messages=e.messages.filter(o=>o.id!==s),e.selectedId=null,e.selectedDetail=null,p()}catch(o){e.error=o instanceof Error?o.message:"Failed to archive message",c()}}else if(a.id==="delete-btn"&&e.selectedDetail){const s=e.selectedDetail.id;try{await le(s),A.delete(s),e.messages=e.messages.filter(o=>o.id!==s),e.selectedId=null,e.selectedDetail=null,p()}catch(o){e.error=o instanceof Error?o.message:"Failed to delete message",c()}}});v.addEventListener("change",async t=>{const a=t.target;if(a.id==="unread-toggle")e.unreadOnly=a.checked,y();else if(a.id==="compose-files"){const n=await V(a.files);e.composeAttachments=[...e.composeAttachments,...n].slice(0,10);const i=document.getElementById("compose-att-pills");i&&(i.innerHTML=T(e.composeAttachments,"compose"))}else if(a.id==="reply-files"){const n=await V(a.files);e.replyAttachments=[...e.replyAttachments,...n].slice(0,10);const i=document.getElementById("reply-att-pills");i&&(i.innerHTML=T(e.replyAttachments,"reply"))}});function E(){const t=document.getElementById("compose-suggestions");if(t){if(f.length===0){t.hidden=!0,t.innerHTML="";return}t.hidden=!1,t.innerHTML=f.map((a,n)=>`
      <button type="button" class="autocomplete-item ${n===b?"active":""}" data-email="${d(a.address)}">
        <span><strong>${d(a.name)}</strong> &lt;${d(a.address)}&gt;</span>
        <span class="autocomplete-sub">${d(a.source)}</span>
      </button>`).join("")}}function Y(t){const a=document.getElementById("compose-to");if(!a)return;const n=a.value.split(",");n[n.length-1]=` ${t}`,a.value=`${n.map(i=>i.trim()).filter(Boolean).join(", ")}, `,f=[],b=-1,E(),a.focus()}v.addEventListener("input",t=>{const a=t.target;if(a.id==="search-input")e.search=a.value,C&&window.clearTimeout(C),C=window.setTimeout(()=>void y(),200);else if(a.id==="compose-to"){const n=a.value.split(",").pop()?.trim()??"";if(M&&window.clearTimeout(M),n.length<1){f=[],b=-1,E();return}M=window.setTimeout(async()=>{D?.abort(),D=new AbortController;try{f=(await pe(n,D.signal)).suggestions,b=f.length>0?0:-1,E()}catch{}},150)}});v.addEventListener("keydown",t=>{t.target.id!=="compose-to"||f.length===0||(t.key==="ArrowDown"?(t.preventDefault(),b=(b+1)%f.length,E()):t.key==="ArrowUp"?(t.preventDefault(),b=(b-1+f.length)%f.length,E()):t.key==="Enter"&&b>=0?(t.preventDefault(),Y(f[b].address)):t.key==="Escape"&&(f=[],b=-1,E()))});v.addEventListener("submit",async t=>{const a=t.target;if(a.id!=="compose-form")return;t.preventDefault();const n=document.getElementById("compose-to")?.value??"",i=document.getElementById("compose-subject")?.value.trim()??"",r=document.getElementById("compose-body")?.value??"",l=n.split(",").map(o=>o.trim()).filter(Boolean);if(l.length===0||!i)return;e.sendingCompose=!0;const s=a.querySelector('button[type="submit"]');s&&(s.disabled=!0,s.textContent="Sending…");try{await me({to:l,subject:i,body:r,attachments:e.composeAttachments.length>0?e.composeAttachments:void 0}),e.composeOpen=!1,e.composeAttachments=[],e.sendingCompose=!1,f=[],b=-1,c(),y()}catch(o){e.sendingCompose=!1,s&&(s.disabled=!1,s.textContent="Send"),e.error=o instanceof Error?o.message:"Failed to send email",c()}});async function K(){if(!(!e.auth?.authenticated||document.hidden||e.composeOpen||e.replyOpen||e.loadingList))try{const t=await _({folder:e.folder,unreadOnly:e.unreadOnly,search:e.search});e.messages=t.messages;const a=document.getElementById("message-list");a&&(a.innerHTML=F())}catch{}}const we=3e4;async function N(){if(window.location.pathname==="/auth-done"){const a=new URLSearchParams(window.location.search).get("sid")??"";try{const{app:n,authentication:i}=await P(async()=>{const{app:r,authentication:l}=await import("./src-CZ2ouLw7.js");return{app:r,authentication:l}},[],import.meta.url);await n.initialize(),i.notifySuccess(a)}catch{window.close()}return}const t=await ge(z);e.inTeams=t.inTeams,z(t.theme),ne(t.ssoToken);try{const a=U(),n=_({folder:e.folder,unreadOnly:e.unreadOnly,search:e.search}).catch(()=>null),[i,r]=await Promise.all([a,n]);e.serverUnreachable=!1,e.auth=i,i.authenticated&&r?e.messages=r.messages:i.authenticated&&!r&&y(),c()}catch(a){if(a instanceof G){e.serverUnreachable=!0,c(),window.setTimeout(()=>{e.serverUnreachable&&N()},4e3);return}e.error=a instanceof Error?a.message:"Failed to initialize",c()}fe(!0).then(a=>{e.versionInfo=a,c()}).catch(()=>{}),window.setInterval(()=>void K(),we),document.addEventListener("visibilitychange",()=>{document.hidden||K()}),document.addEventListener("keydown",a=>{(a.metaKey||a.ctrlKey)&&a.key==="Enter"?e.composeOpen?(a.preventDefault(),document.getElementById("compose-form")?.requestSubmit()):e.replyOpen?(a.preventDefault(),document.getElementById("send-reply-btn")?.click()):e.forwardOpen&&(a.preventDefault(),document.getElementById("send-forward-btn")?.click()):a.key==="Escape"&&(e.composeOpen?(e.composeOpen=!1,e.composeAttachments=[],c()):(e.replyOpen||e.forwardOpen)&&(e.replyOpen=!1,e.replyAttachments=[],e.forwardOpen=!1,p()))})}N();
