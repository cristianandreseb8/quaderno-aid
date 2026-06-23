import React, { useState, useEffect, useRef, useMemo, useCallback } from 'react'
import { supabase } from './lib/supabase.js'
import jsPDF from 'jspdf'
import * as XLSX from 'xlsx'

/* ═══════════════════════════════════════════════════════════════
   QUADERNO AI+D  v1.0
   Base: Quaderno AI v2 (all features preserved)
   New:
   · Fix: thumbnail preserved in PDF/Image after translation
   · Copy recipe → fixed language version
   · Módulo I+D (pestaña separada por receta):
       - Ficha de parámetros (grasas, agua, azúcar, sal, hidratación…)
       - Fichas de percepción sensorial + sistema de puntaje
       - Timeline de versiones (evolución hacia resultado final)
       - Comparativa entre recetas
       - Tabla nutricional básica
   · Biblioteca de medios por receta y en notas (fotos, audio ≤2 min, video ≤3 min)
   · Voz con corrección inteligente de términos técnicos de panadería
   ═══════════════════════════════════════════════════════════════ */

const CSS = `
.Q{--paper:#FAF7F0;--ink:#221C18;--muted:#6E645C;--rule:#E6DECF;
  --navy:#1F3A4D;--amber:#BC6C2C;--warm:#FBEFE1;--green:#2D6A4F;--ai:#5B3A8C;
  --red:#9b2c2c;--teal:#1A6B6B;--id:#1A5276;
  --serif:Georgia,"Iowan Old Style",serif;
  --sans:ui-sans-serif,system-ui,-apple-system,"Segoe UI",sans-serif;
  --mono:ui-monospace,"SF Mono","Menlo",monospace;
  font-family:var(--sans);color:var(--ink);background:var(--paper);min-height:100vh;display:flex;flex-direction:column}
.Q *{box-sizing:border-box}
.Q-top{display:flex;align-items:center;gap:10px;flex-wrap:wrap;padding:12px 20px;
  border-bottom:1px solid var(--rule);background:var(--paper);position:sticky;top:0;z-index:20}
.Q-brand{font-family:var(--serif);font-size:21px;color:var(--navy);white-space:nowrap;display:flex;align-items:baseline;gap:6px}
.Q-brand .ai-badge{font-family:var(--mono);font-size:9px;color:#fff;background:var(--ai);padding:2px 6px;border-radius:10px;letter-spacing:.1em;text-transform:uppercase}
.Q-brand .id-badge{font-family:var(--mono);font-size:9px;color:#fff;background:var(--id);padding:2px 6px;border-radius:10px;letter-spacing:.1em;text-transform:uppercase;margin-left:2px}
.Q-top-right{margin-left:auto;display:flex;gap:7px;align-items:center}
.btn{font-family:var(--sans);font-size:12px;font-weight:600;cursor:pointer;border:1.5px solid var(--navy);background:transparent;color:var(--navy);padding:6px 11px;border-radius:6px;transition:background .13s,color .13s;white-space:nowrap}
.btn:hover{background:var(--navy);color:var(--paper)}
.btn:focus-visible{outline:2px solid var(--amber);outline-offset:2px}
.btn:disabled{opacity:.4;cursor:default}
.btn.amber{border-color:var(--amber);color:var(--amber)}.btn.amber:hover{background:var(--amber);color:#fff}
.btn.ghost{border-color:var(--rule);color:var(--muted)}.btn.ghost:hover{background:#ede6d9;color:var(--ink)}
.btn.green{border-color:var(--green);color:var(--green)}.btn.green:hover{background:var(--green);color:#fff}
.btn.ai{border-color:var(--ai);color:var(--ai)}.btn.ai:hover{background:var(--ai);color:#fff}
.btn.id{border-color:var(--id);color:var(--id)}.btn.id:hover{background:var(--id);color:#fff}
.btn.danger{border-color:var(--red);color:var(--red)}.btn.danger:hover{background:var(--red);color:#fff}
.btn.teal{border-color:var(--teal);color:var(--teal)}.btn.teal:hover{background:var(--teal);color:#fff}
.btn.xs{font-size:10.5px;padding:4px 8px}
.Q-body{flex:1;display:grid;grid-template-columns:282px 1fr;min-height:0}
.Q-side{border-right:1px solid var(--rule);display:flex;flex-direction:column;min-height:0}
.Q-search{padding:10px 12px;border-bottom:1px solid var(--rule)}
.Q-search input{width:100%;border:1px solid var(--rule);background:#fff;border-radius:6px;padding:7px 10px;font-size:12.5px;font-family:var(--sans);color:var(--ink)}
.Q-search input:focus{outline:none;border-color:var(--navy)}
.Q-list{overflow:auto;flex:1}
.Q-list-item{width:100%;text-align:left;background:none;border:none;cursor:pointer;padding:9px 13px;border-bottom:1px solid var(--rule);display:flex;align-items:center;gap:8px}
.Q-list-item:hover{background:#f3ede0}
.Q-list-item[aria-selected=true]{background:#fff;box-shadow:inset 3px 0 0 var(--amber)}
.Q-list-thumb{width:34px;height:34px;border-radius:5px;object-fit:cover;flex-shrink:0}
.Q-list-thumb-ph{width:34px;height:34px;border-radius:5px;background:var(--rule);flex-shrink:0;display:flex;align-items:center;justify-content:center;font-size:13px}
.Q-list-item h4{font-family:var(--serif);font-size:14px;font-weight:600;margin:0 0 2px}
.Q-list-item span{font-family:var(--mono);font-size:9.5px;color:var(--muted);text-transform:uppercase;letter-spacing:.1em}
.Q-msg{padding:20px 14px;color:var(--muted);font-size:12.5px;line-height:1.6}
.Q-main{overflow:auto}
.Q-pane{max-width:820px;margin:0 auto;padding:28px 36px 100px}
.Q-tabs{display:flex;border-bottom:2px solid var(--rule);margin:14px 0 18px;gap:0;flex-wrap:wrap}
.Q-tab-btn{font-family:var(--sans);font-size:12.5px;font-weight:600;background:none;border:none;padding:8px 14px;cursor:pointer;color:var(--muted);border-bottom:2px solid transparent;margin-bottom:-2px;transition:color .13s;white-space:nowrap}
.Q-tab-btn:hover{color:var(--ink)}
.Q-tab-btn.active{color:var(--navy);border-bottom-color:var(--amber)}
.Q-tab-btn.ai-tab.active{color:var(--ai);border-bottom-color:var(--ai)}
.Q-tab-btn.id-tab.active{color:var(--id);border-bottom-color:var(--id)}
.Q-view{animation:vfade .25s ease both}
@keyframes vfade{from{opacity:0;transform:translateY(4px)}to{opacity:1;transform:none}}
.Q-view-header{display:flex;gap:14px;align-items:flex-start;margin-bottom:10px}
.Q-view-header h1{font-family:var(--serif);font-size:28px;line-height:1.18;margin:0;flex:1}
.Q-recipe-thumb{width:80px;height:80px;border-radius:8px;object-fit:cover;cursor:pointer;flex-shrink:0;border:1px solid var(--rule)}
.Q-recipe-thumb:hover{opacity:.85}
.Q-meta{display:flex;flex-wrap:wrap;gap:14px;padding-bottom:12px;border-bottom:2px solid var(--ink);margin-bottom:14px}
.Q-meta-item dt{font-family:var(--mono);font-size:9px;text-transform:uppercase;letter-spacing:.14em;color:var(--muted);margin-bottom:2px}
.Q-meta-item dd{margin:0;font-size:13px;font-weight:600}
.Q-banner{display:flex;align-items:center;gap:10px;padding:7px 11px;border-radius:7px;font-size:11.5px;margin-bottom:10px;flex-wrap:wrap}
.Q-banner.scale{background:#EAF2EE;border:1px solid #a8d5bc;color:var(--green)}
.Q-banner.trans{background:#EEF1F5;border:1px solid #b8c8d8;color:var(--navy)}
.Q-banner.copy{background:#FFF3E0;border:1px solid #FFB74D;color:#E65100}
.Q-banner button{margin-left:auto;font-size:11px;font-family:var(--mono);background:none;border:none;cursor:pointer;text-decoration:underline;color:inherit}
.Q-toolbar{display:flex;gap:5px;flex-wrap:wrap;margin-bottom:14px;align-items:center}
.Q-toolbar .right{margin-left:auto;display:flex;gap:5px}
.Q-export-opts{display:flex;align-items:center;gap:8px;padding:7px 11px;background:#f5f0e8;border-radius:7px;margin-bottom:12px;font-size:12px;color:var(--muted)}
.Q-export-opts label{display:flex;align-items:center;gap:5px;cursor:pointer;color:var(--ink)}
.Q-scale-panel{background:#fff;border:1px solid var(--rule);border-radius:9px;padding:14px 16px;margin-bottom:14px;display:flex;flex-direction:column;gap:10px}
.Q-scale-panel h4{font-family:var(--mono);font-size:10px;text-transform:uppercase;letter-spacing:.16em;color:var(--navy);margin:0 0 2px}
.Q-scale-row{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.Q-scale-row label{font-size:12.5px;color:var(--muted);white-space:nowrap}
.Q-scale-row input{border:1px solid var(--rule);border-radius:5px;padding:5px 8px;width:80px;font-size:13px;font-family:var(--mono);color:var(--ink)}
.Q-scale-row input:focus{outline:none;border-color:var(--navy)}
.Q-pct-bar{display:flex;gap:6px;align-items:center;flex-wrap:wrap;margin-bottom:8px}
.Q-pct-bar label{font-family:var(--mono);font-size:9.5px;text-transform:uppercase;letter-spacing:.14em;color:var(--navy)}
.Q-pct-bar select,.Q-scale-row select{border:1px solid var(--rule);border-radius:5px;padding:4px 6px;font-size:11.5px;font-family:var(--mono);color:var(--ink);background:#fff}
.Q-sec-h{display:flex;align-items:center;gap:8px;margin:18px 0 7px;font-family:var(--mono);font-size:10px;text-transform:uppercase;letter-spacing:.18em;color:var(--navy)}
.Q-sec-h::before,.Q-sec-h::after{content:'';flex:1;height:1px;background:var(--rule)}
.Q-ings{list-style:none;padding:0;margin:0 0 3px}
.Q-ing-row{display:flex;align-items:baseline;padding:5px 0;border-bottom:1px dotted var(--rule);cursor:pointer;border-radius:4px;transition:background .12s}
.Q-ing-row:hover{background:#f5efea}
.Q-ing-check{width:22px;min-width:22px;height:18px;display:flex;align-items:center;justify-content:center;font-size:11px;color:var(--muted)}
.Q-ing-row.checked{background:#FCEBD9}
.Q-ing-row.checked .Q-ing-check{color:var(--amber)}
.Q-ing-row.checked .Q-ing-qty,.Q-ing-row.checked .Q-ing-name{color:var(--amber);text-decoration:line-through;text-decoration-color:rgba(188,108,44,.4)}
.Q-ing-qty{font-family:var(--mono);font-size:12px;color:var(--muted);min-width:86px;text-align:right;white-space:nowrap;padding-right:10px}
.Q-ing-name{font-size:13px;flex:1}
.Q-pct-badge{font-family:var(--mono);font-size:10px;color:var(--muted);margin-left:6px;white-space:nowrap;min-width:44px;text-align:right}
.Q-pct-badge.base{color:var(--amber);font-weight:700}
.Q-subtotal{font-family:var(--mono);font-size:10px;color:var(--muted);text-align:right;padding:4px 0;letter-spacing:.04em;border-top:1px solid var(--rule)}
.Q-grand-total{font-family:var(--mono);font-size:11px;font-weight:700;color:var(--navy);text-align:right;padding:7px 0;border-top:2px solid var(--navy);margin-bottom:22px}
.Q-steps-label{font-family:var(--mono);font-size:10px;text-transform:uppercase;letter-spacing:.18em;color:var(--navy);margin:0 0 9px}
.Q-steps{list-style:none;padding:0;margin:0;counter-reset:s}
.Q-steps li{counter-increment:s;position:relative;padding:6px 10px 11px 40px;font-size:13px;line-height:1.6;border-radius:6px;transition:background .15s}
.Q-steps li::before{content:counter(s,decimal-leading-zero);position:absolute;left:0;top:7px;font-family:var(--mono);font-size:10.5px;color:var(--amber);font-weight:700;width:36px;text-align:right}
.Q-steps li.highlighted{background:#FCEBD9;border-left:3px solid var(--amber)}
.Q-steps li.highlighted::before{color:var(--navy)}
.Q-baker-note{margin-top:16px;padding:11px 13px;background:#fff;border:1px solid var(--rule);border-left:3.5px solid var(--amber);border-radius:7px;font-size:12.5px;line-height:1.55;color:var(--muted)}

/* drag-drop */

.Q-drag-list{border:1px solid var(--rule);border-radius:7px;background:#fff;overflow:hidden}
.Q-drag-item{display:flex;align-items:center;gap:6px;padding:6px 10px;border-bottom:1px solid var(--rule);background:#fff;transition:background .1s}
.Q-drag-item.over{background:#f0ebe2;border-left:2px solid var(--amber)}
.Q-drag-item.dragging{opacity:.3}
.Q-drag-item.is-section{background:#FAF7F0}
.Q-drag-handle{cursor:grab;color:var(--rule);font-size:15px;flex-shrink:0;user-select:none;line-height:1}
.Q-drag-handle:active{cursor:grabbing}
.Q-drag-input{flex:1;border:none;background:transparent;font-size:13px;font-family:var(--mono);color:var(--ink);outline:none;padding:2px 0}
.Q-drag-input.section{font-size:10.5px;text-transform:uppercase;letter-spacing:.14em;color:var(--amber);font-weight:700}
.Q-drag-rm{background:none;border:none;cursor:pointer;color:var(--rule);font-size:13px;padding:0 2px;flex-shrink:0}
.Q-drag-rm:hover{color:var(--red)}
.Q-drag-footer{display:flex;gap:6px;padding:8px 10px;background:#f5f0e8}

/* notes tabs */

.Q-notes-panel{display:flex;flex-direction:column}
.Q-notes-tabs-row{display:flex;align-items:center;gap:0;border-bottom:1px solid var(--rule);flex-wrap:wrap}
.Q-notes-tab{font-size:12px;font-family:var(--sans);font-weight:600;background:none;border:none;padding:7px 12px;cursor:pointer;color:var(--muted);border-bottom:2px solid transparent;margin-bottom:-1px;white-space:nowrap}
.Q-notes-tab.active{color:var(--navy);border-bottom-color:var(--amber)}
.Q-notes-tab-add{background:none;border:none;cursor:pointer;color:var(--muted);font-size:16px;padding:4px 8px;line-height:1}
.Q-notes-tab-add:hover{color:var(--amber)}
.Q-notes-toolbar{display:flex;align-items:center;justify-content:space-between;padding:8px 0 6px;flex-wrap:wrap;gap:6px}
.Q-notes-textarea{width:100%;min-height:280px;border:1px solid var(--rule);border-radius:8px;padding:13px;font-size:13.5px;line-height:1.7;font-family:var(--sans);color:var(--ink);resize:vertical;background:#fff}
.Q-notes-textarea:focus{outline:none;border-color:var(--navy)}
.Q-recording-pill{display:inline-flex;align-items:center;gap:5px;background:#fee;border:1px solid #f88;border-radius:20px;padding:3px 9px;font-size:11px;color:#a23b2e;animation:blink 1.2s infinite}
@keyframes blink{0%,100%{opacity:1}50%{opacity:.5}}
.Q-voice-btn{background:none;border:1.5px solid var(--rule);border-radius:50%;width:30px;height:30px;cursor:pointer;font-size:13px;display:flex;align-items:center;justify-content:center;transition:all .13s;flex-shrink:0}
.Q-voice-btn:hover{border-color:var(--amber);background:var(--warm)}
.Q-voice-btn.recording{border-color:#c0392b;background:#fee;animation:blink 1s infinite}

/* media library */

.Q-media-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(120px,1fr));gap:10px;margin:12px 0}
.Q-media-card{position:relative;border:1px solid var(--rule);border-radius:8px;overflow:hidden;background:#fff;cursor:pointer;transition:box-shadow .15s}
.Q-media-card:hover{box-shadow:0 2px 8px rgba(0,0,0,.12)}
.Q-media-thumb{width:100%;height:90px;object-fit:cover;display:block}
.Q-media-audio{display:flex;flex-direction:column;align-items:center;justify-content:center;height:90px;background:#f0ebe2;gap:4px}
.Q-media-video{display:flex;flex-direction:column;align-items:center;justify-content:center;height:90px;background:#1a2a3a;gap:4px;color:#fff}
.Q-media-label{font-family:var(--mono);font-size:9px;text-transform:uppercase;letter-spacing:.1em;color:var(--muted);padding:4px 6px;text-align:center;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.Q-media-rm{position:absolute;top:4px;right:4px;width:20px;height:20px;background:rgba(0,0,0,.65);color:#fff;border:none;border-radius:50%;cursor:pointer;font-size:11px;display:flex;align-items:center;justify-content:center;opacity:0;transition:opacity .15s}
.Q-media-card:hover .Q-media-rm{opacity:1}
.Q-media-player-wrap{background:rgba(0,0,0,.92);position:fixed;inset:0;z-index:400;display:flex;align-items:center;justify-content:center;flex-direction:column;gap:12px}
.Q-media-player-wrap video,.Q-media-player-wrap audio{max-width:90vw;max-height:80vh;border-radius:8px}
.Q-media-upload-btn{display:flex;align-items:center;gap:7px;padding:10px 14px;border:1.5px dashed var(--amber);border-radius:8px;background:rgba(188,108,44,.04);cursor:pointer;font-size:12.5px;color:var(--amber);font-weight:600;transition:background .15s}
.Q-media-upload-btn:hover{background:rgba(188,108,44,.1)}

/* AI assistant */

.Q-assistant{display:flex;flex-direction:column;min-height:400px}
.Q-assistant-welcome{text-align:center;padding:24px 16px;color:var(--muted)}
.Q-chat-wrap{flex:1;overflow:auto;padding:6px 0;display:flex;flex-direction:column;gap:9px;max-height:380px}
.Q-chat-msg{max-width:86%;padding:9px 12px;border-radius:10px;font-size:13px;line-height:1.55}
.Q-chat-msg.user{align-self:flex-end;background:var(--navy);color:#fff;border-bottom-right-radius:3px}
.Q-chat-msg.assistant{align-self:flex-start;background:#fff;border:1px solid var(--rule);border-bottom-left-radius:3px;color:var(--ink)}
.Q-chat-action-badge{margin-top:5px;font-family:var(--mono);font-size:10px;color:var(--green)}
.Q-chat-typing{display:flex;gap:4px;padding:2px 0}
.Q-chat-typing span{width:7px;height:7px;border-radius:50%;background:var(--muted);animation:typing .8s infinite}
.Q-chat-typing span:nth-child(2){animation-delay:.15s}
.Q-chat-typing span:nth-child(3){animation-delay:.3s}
@keyframes typing{0%,60%,100%{transform:translateY(0)}30%{transform:translateY(-5px)}}
.Q-conv-toolbar{display:flex;align-items:center;gap:8px;padding:6px 0;border-bottom:1px solid var(--rule);margin-bottom:8px;flex-wrap:wrap;font-size:11.5px;color:var(--muted)}
.Q-chat-input-area{border-top:1px solid var(--rule);padding-top:9px;display:flex;gap:7px;align-items:flex-end;margin-top:8px}
.Q-chat-input{flex:1;border:1px solid var(--rule);border-radius:8px;padding:8px 10px;font-size:13px;font-family:var(--sans);color:var(--ink);resize:none;background:#fff}
.Q-chat-input:focus{outline:none;border-color:var(--ai)}
.Q-quick-chips{display:flex;flex-wrap:wrap;gap:4px;margin-bottom:10px}
.Q-chip{font-size:11px;padding:3px 9px;border:1px solid var(--rule);border-radius:20px;background:#fff;cursor:pointer;color:var(--muted);transition:all .12s;font-family:var(--mono)}
.Q-chip:hover{border-color:var(--ai);color:var(--ai);background:#F5F0FF}

/* App AI panel */

.Q-app-ai-overlay{position:fixed;inset:0;z-index:200;background:rgba(0,0,0,.15)}
.Q-app-ai-panel{position:absolute;right:0;top:0;bottom:0;width:480px;max-width:100vw;background:var(--paper);border-left:1px solid var(--rule);display:flex;flex-direction:column;box-shadow:-4px 0 24px rgba(0,0,0,.1)}
.Q-app-ai-header{display:flex;align-items:center;padding:14px 16px;border-bottom:1px solid var(--rule)}
.Q-app-ai-title{font-family:var(--serif);font-size:17px;color:var(--navy);flex:1}
.Q-app-ai-msgs{flex:1;overflow:auto;padding:12px 16px;display:flex;flex-direction:column;gap:10px}
.Q-app-ai-input{border-top:1px solid var(--rule);padding:12px 16px;display:flex;gap:7px;align-items:flex-end}

/* lightbox */

.Q-lightbox{position:fixed;inset:0;background:rgba(0,0,0,.9);z-index:300;display:flex;align-items:center;justify-content:center;cursor:pointer}
.Q-lightbox img{max-width:90vw;max-height:90vh;object-fit:contain;border-radius:8px}

/* hero */

.Q-hero{max-width:500px;margin:10vh auto 0;text-align:center;padding:20px}
.Q-hero .glyph{font-family:var(--serif);font-size:60px;color:var(--amber);line-height:1}
.Q-hero h2{font-family:var(--serif);font-size:22px;margin:14px 0 7px}
.Q-hero p{color:var(--muted);font-size:13px;line-height:1.65;margin:0 auto 20px;max-width:360px}

/* editor */

.Q-ed h2{font-family:var(--serif);font-size:22px;margin:0 0 20px}
.Q-field{margin-bottom:13px}
.Q-field label{display:block;font-family:var(--mono);font-size:9.5px;text-transform:uppercase;letter-spacing:.14em;color:var(--muted);margin-bottom:4px}
.Q-field input,.Q-field textarea{width:100%;border:1px solid var(--rule);background:#fff;border-radius:6px;padding:8px 10px;font-size:13px;font-family:var(--sans);color:var(--ink);resize:vertical}
.Q-field input:focus,.Q-field textarea:focus{outline:none;border-color:var(--navy)}
.Q-field .hint{font-size:10.5px;color:var(--muted);margin-top:4px;line-height:1.5}
.Q-grid2{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.Q-ed-foot{display:flex;gap:7px;margin-top:20px;flex-wrap:wrap}
.Q-thumbs{display:flex;gap:7px;flex-wrap:wrap;margin:9px 0 3px}
.Q-thumb{position:relative;width:62px;height:62px;border-radius:5px;overflow:hidden;border:1px solid var(--rule)}
.Q-thumb img{width:100%;height:100%;object-fit:cover;display:block}
.Q-thumb button{position:absolute;top:2px;right:2px;width:16px;height:16px;border:none;border-radius:50%;background:rgba(0,0,0,.65);color:#fff;cursor:pointer;font-size:10px;line-height:1;padding:0}
.Q-err{color:var(--red);font-size:11.5px;margin-top:7px;line-height:1.5}

/* I+D Module */

.ID-panel{display:flex;flex-direction:column;gap:0}
.ID-section{background:#fff;border:1px solid var(--rule);border-radius:10px;margin-bottom:14px;overflow:hidden}
.ID-section-header{display:flex;align-items:center;gap:8px;padding:11px 14px;border-bottom:1px solid var(--rule);background:#F0F4F8;cursor:pointer;user-select:none}
.ID-section-header h3{font-family:var(--mono);font-size:10px;text-transform:uppercase;letter-spacing:.18em;color:var(--id);margin:0;flex:1}
.ID-section-body{padding:14px}
.ID-param-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:10px;margin-bottom:8px}
.ID-param-card{background:#F8FBFF;border:1px solid #C8DFF0;border-radius:7px;padding:10px 12px}
.ID-param-label{font-family:var(--mono);font-size:8.5px;text-transform:uppercase;letter-spacing:.14em;color:var(--id);margin-bottom:4px}
.ID-param-value{font-family:var(--mono);font-size:18px;font-weight:700;color:var(--navy)}
.ID-param-unit{font-size:10px;color:var(--muted);margin-left:3px;font-weight:400}
.ID-param-bar{height:4px;border-radius:2px;background:var(--rule);margin-top:5px;overflow:hidden}
.ID-param-bar-fill{height:100%;border-radius:2px;background:var(--id);transition:width .4s}

/* Sensory */

.ID-sensory-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:10px}
.ID-sensory-item{background:#F8FBFF;border:1px solid #C8DFF0;border-radius:7px;padding:10px}
.ID-sensory-label{font-family:var(--mono);font-size:9px;text-transform:uppercase;letter-spacing:.12em;color:var(--id);margin-bottom:6px}
.ID-star-row{display:flex;gap:3px;margin-bottom:4px}
.ID-star{font-size:18px;cursor:pointer;transition:transform .1s;line-height:1}
.ID-star:hover{transform:scale(1.2)}
.ID-sensory-note{width:100%;border:1px solid var(--rule);border-radius:5px;padding:5px 7px;font-size:12px;font-family:var(--sans);color:var(--ink);resize:none}
.ID-sensory-note:focus{outline:none;border-color:var(--id)}

/* Timeline */

.ID-timeline{position:relative;padding-left:24px}
.ID-timeline::before{content:'';position:absolute;left:8px;top:0;bottom:0;width:2px;background:var(--rule)}
.ID-timeline-item{position:relative;margin-bottom:16px}
.ID-timeline-dot{position:absolute;left:-20px;top:4px;width:12px;height:12px;border-radius:50%;background:var(--id);border:2px solid #fff;box-shadow:0 0 0 2px var(--id)}
.ID-timeline-dot.active{background:var(--amber);box-shadow:0 0 0 2px var(--amber)}
.ID-timeline-dot.goal{background:var(--green);box-shadow:0 0 0 2px var(--green)}
.ID-timeline-card{background:#F8FBFF;border:1px solid #C8DFF0;border-radius:7px;padding:10px 12px}
.ID-timeline-date{font-family:var(--mono);font-size:9px;color:var(--muted);margin-bottom:3px}
.ID-timeline-title{font-size:13px;font-weight:600;margin-bottom:4px}
.ID-timeline-notes{font-size:12px;color:var(--muted);line-height:1.5}
.ID-timeline-badge{display:inline-block;font-family:var(--mono);font-size:8.5px;text-transform:uppercase;letter-spacing:.1em;padding:2px 7px;border-radius:10px;margin-bottom:5px}
.ID-timeline-badge.goal{background:#EAF2EE;color:var(--green)}
.ID-timeline-badge.version{background:#EEF1F5;color:var(--id)}

/* Nutrition */

.ID-nutrition-table{width:100%;border-collapse:collapse;font-size:12.5px}
.ID-nutrition-table th{font-family:var(--mono);font-size:9px;text-transform:uppercase;letter-spacing:.12em;color:var(--id);text-align:left;padding:6px 10px;border-bottom:2px solid var(--id)}
.ID-nutrition-table td{padding:6px 10px;border-bottom:1px solid var(--rule)}
.ID-nutrition-table tr:nth-child(even) td{background:#F8FBFF}

/* Compare */

.ID-compare-bar{height:18px;border-radius:4px;background:var(--rule);position:relative;overflow:hidden;margin:3px 0}
.ID-compare-fill{height:100%;border-radius:4px;transition:width .5s}
.ID-compare-label{font-size:10px;font-family:var(--mono);position:absolute;right:6px;top:2px;color:#fff;font-weight:700;text-shadow:0 1px 2px rgba(0,0,0,.4)}

/* Compare panel */

.Q-compare-overlay{position:fixed;inset:0;z-index:250;background:rgba(0,0,0,.18)}
.Q-compare-panel{position:absolute;right:0;top:0;bottom:0;width:600px;max-width:100vw;background:var(--paper);border-left:1px solid var(--rule);display:flex;flex-direction:column;box-shadow:-4px 0 24px rgba(0,0,0,.1)}
.Q-compare-header{display:flex;align-items:center;padding:14px 16px;border-bottom:1px solid var(--rule)}
.Q-compare-body{flex:1;overflow:auto;padding:16px}

/* responsive */

.Q-back-btn{display:none}
@media(max-width:700px){
  .Q-body{grid-template-columns:1fr}
  .Q-side{display:flex}.Q-main{display:none}
  .Q[data-open="1"] .Q-side{display:none}
  .Q[data-open="1"] .Q-main{display:block}
  .Q-back-btn{display:inline-flex}
  .Q-pane{padding:16px 14px 80px}
  .Q-toolbar .right{margin-left:0}
  .Q-app-ai-panel,.Q-compare-panel{width:100vw}
  .ID-param-grid{grid-template-columns:1fr 1fr}
}
@media(prefers-reduced-motion:reduce){.Q-view,.Q-chat-typing span,.Q-recording-pill{animation:none}}
`

/* ── Utilities ─────────────────────────────────────────────────────────────── */

const uid = () => Math.random().toString(36).slice(2,9)+Date.now().toString(36)
const ts = () => new Date().toLocaleString('en-US',{month:'short',day:'numeric',year:'numeric',hour:'2-digit',minute:'2-digit'})
function parseIng(text){const t=String(text||'').trim();const m=t.match(/^([\d.,]+(?:\/[\d.,]+)?)\s*([a-zA-Z%]*)\s{1,}(.+)$/);if(!m)return{qty:null,unit:'',name:t};const qty=m[1].includes('/')?m[1].split('/').reduce((a,b)=>parseFloat(a)/parseFloat(b)):parseFloat(m[1].replace(',','.'));return{qty,unit:m[2].toLowerCase(),name:m[3].trim()}}
function toGrams(qty,unit){if(!qty||isNaN(qty))return 0;const u=unit||'';if(u==='kg')return qty*1000;if(u==='l')return qty*1000;if(u==='ml')return qty;if(u==='%')return 0;return qty}
const FLOUR_W=['flour','farina','harina','mehl','farine','semolina','semola','manitoba','grano','t45','t55','t65','t80','t150','00','tipo']
function isFlour(n){const s=(n||'').toLowerCase();return FLOUR_W.some(k=>s.includes(k))}
function fmtQty(q){if(q>=100)return String(Math.round(q));if(q>=10)return(Math.round(q*10)/10).toFixed(1);return(Math.round(q*100)/100).toFixed(q<1?2:1)}
function parseSections(ingredients){const ings=ingredients||[];const sections=[];let cur={name:null,items:[],rawIndices:[]};ings.forEach((ing,i)=>{if(/^##?\s+/.test(ing)){if(cur.items.length||cur.name!==null)sections.push(cur);cur={name:ing.replace(/^##?\s*/,'').trim(),items:[],rawIndices:[]}}else{cur.items.push(ing);cur.rawIndices.push(i)}});if(cur.items.length||cur.name!==null)sections.push(cur);if(!sections.length)return[{name:null,items:ings,rawIndices:ings.map((_,i)=>i)}];return sections}
function calcPct(items,mode,base){const parsed=items.map(i=>{const p=parseIng(i);return{...p,grams:toGrams(p.qty,p.unit)}});let bg=0;if(mode==='baker')bg=parsed.filter(p=>isFlour(p.name)).reduce((s,p)=>s+p.grams,0);else if(mode==='mass')bg=parsed.reduce((s,p)=>s+p.grams,0);else if(mode==='custom'&&base){const b=parsed.find(p=>p.name.toLowerCase().includes(base.toLowerCase()));bg=b?b.grams:0};return parsed.map(p=>({...p,pct:bg>0&&p.grams>0?p.grams/bg*100:null,isBase:mode==='custom'&&base&&p.name.toLowerCase().includes(base.toLowerCase())}))}
function getTotalGrams(ingredients){return(ingredients||[]).reduce((s,ing)=>{if(/^##?\s+/.test(ing))return s;const p=parseIng(ing);return s+toGrams(p.qty,p.unit)},0)}
function scaleRecipe(recipe,factor){return{...recipe,ingredients:(recipe.ingredients||[]).map(ing=>{if(/^##?\s+/.test(ing))return ing;const p=parseIng(ing);if(p.qty===null)return ing;return `${fmtQty(p.qty*factor)}${p.unit?' '+p.unit:''}  ${p.name}`})}}
function findStepsForIng(name,steps){const words=name.toLowerCase().split(/\s+/).filter(w=>w.length>3);if(!words.length)return new Set();const r=new Set();(steps||[]).forEach((s,i)=>{if(words.some(w=>s.toLowerCase().includes(w)))r.add(i)});return r}

/* ── Macro Analysis ──────────────────────────────────────────────────────────── */

// Approximate nutritional factors per 100g of ingredient type

const NUTRIENT_DB = {
  flour:    {fat:1.2, water:14,  sugar:1.5,  protein:12,  carbs:72,  cal:340,  flourEq:100, freeWater:14},
  butter:   {fat:80,  water:16,  sugar:0,    protein:1,   carbs:0,   cal:720,  flourEq:0,   freeWater:0},
  egg:      {fat:10,  water:74,  sugar:0.4,  protein:13,  carbs:1,   cal:155,  flourEq:0,   freeWater:20},
  sugar:    {fat:0,   water:0,   sugar:100,  protein:0,   carbs:100, cal:387,  flourEq:0,   freeWater:0},
  milk:     {fat:3.5, water:88,  sugar:5,    protein:3.4, carbs:5,   cal:61,   flourEq:0,   freeWater:88},
  cream:    {fat:35,  water:58,  sugar:3,    protein:2.5, carbs:3,   cal:340,  flourEq:0,   freeWater:30},
  salt:     {fat:0,   water:0,   sugar:0,    protein:0,   carbs:0,   cal:0,    flourEq:0,   freeWater:0},
  yeast:    {fat:1.5, water:70,  sugar:0,    protein:40,  carbs:18,  cal:105,  flourEq:0,   freeWater:70},
  honey:    {fat:0,   water:17,  sugar:82,   protein:0.3, carbs:82,  cal:304,  flourEq:0,   freeWater:10},
  chocolate:{fat:32,  water:1,   sugar:48,   protein:8,   carbs:55,  cal:546,  flourEq:0,   freeWater:0},
  oil:      {fat:100, water:0,   sugar:0,    protein:0,   carbs:0,   cal:884,  flourEq:0,   freeWater:0},
  water:    {fat:0,   water:100, sugar:0,    protein:0,   carbs:0,   cal:0,    flourEq:0,   freeWater:100},
  egg_yolk: {fat:27,  water:49,  sugar:0.5,  protein:16,  carbs:1,   cal:322,  flourEq:0,   freeWater:20},
  molasses: {fat:0,   water:22,  sugar:75,   protein:0,   carbs:75,  cal:290,  flourEq:0,   freeWater:15},
  sourdough:{fat:1,   water:44,  sugar:0.5,  protein:7,   carbs:45,  cal:220,  flourEq:50,  freeWater:0},
  biga:     {fat:1,   water:42,  sugar:0.5,  protein:8,   carbs:46,  cal:225,  flourEq:60,  freeWater:0},
  poolish:  {fat:1,   water:50,  sugar:0.5,  protein:7,   carbs:44,  cal:215,  flourEq:50,  freeWater:0}
}
function detectIngType(name){
  const s=(name||'').toLowerCase()
  if(FLOUR_W.some(k=>s.includes(k)))return 'flour'
  if(['butter','beurre','mantequilla','burro','margarine'].some(k=>s.includes(k)))return 'butter'
  if(['tuorlo','yolk','jaune'].some(k=>s.includes(k)))return 'egg_yolk'
  if(['egg','uovo','huevo','oeuf'].some(k=>s.includes(k)))return 'egg'
  if(['sugar','sucre','azucar','zucchero','zucker'].some(k=>s.includes(k)))return 'sugar'
  if(['molasses','molasse','melassa','melaza','treacle'].some(k=>s.includes(k)))return 'molasses'
  if(['honey','miel','miele','honig'].some(k=>s.includes(k)))return 'honey'
  if(['milk','lait','leche','latte','milch'].some(k=>s.includes(k)))return 'milk'
  if(['cream','creme','crema','sahne'].some(k=>s.includes(k)))return 'cream'
  if(['salt','sel','sal','sale'].some(k=>s.includes(k)))return 'salt'
  if(['pasta madre','lievito madre','pms','pm s','pm solid','levain','sourdough','poolish','biga','starter','masa madre'].some(k=>s.includes(k))){
    if(s.includes('poolish'))return 'poolish'
    if(s.includes('biga'))return 'biga'
    return 'sourdough'
  }
  if(['oil','olio','aceite','huile'].some(k=>s.includes(k)))return 'oil'
  if(['water','agua','acqua','eau'].some(k=>s.includes(k)))return 'water'
  if(['chocolate','cacao','cocoa'].some(k=>s.includes(k)))return 'chocolate'
  if(['yeast','levure','levadura','lievito'].some(k=>s.includes(k)))return 'yeast'
  return null
}
function calcMacros(ingredients, aiCache){
  const items=(ingredients||[]).filter(i=>!/^##?\s+/.test(i))
  let fat=0,water=0,sugar=0,protein=0,carbs=0,cal=0,total=0,saltG=0,flourEqG=0,freeWaterG=0
  items.forEach(ing=>{
    const p=parseIng(ing);const g=toGrams(p.qty,p.unit)
    if(!g)return
    total+=g
    const cached=aiCache&&aiCache[p.name]
    const t=detectIngType(p.name)
    const db=cached||NUTRIENT_DB[t]||{}
    const fv=(db.fat_pct!=null?db.fat_pct:db.fat)||0
    const wv=(db.water_pct!=null?db.water_pct:db.water)||0
    const fw=(db.free_water_pct!=null?db.free_water_pct:(db.freeWater!=null?db.freeWater:wv))||0
    const fe=(db.flour_equivalent_pct!=null?db.flour_equivalent_pct:(db.flourEq!=null?db.flourEq:0))||0
    const sv=(db.sugar_pct!=null?db.sugar_pct:db.sugar)||0
    const pv=(db.protein_pct!=null?db.protein_pct:db.protein)||0
    const cv=(db.carbs_pct!=null?db.carbs_pct:db.carbs)||0
    const kv=(db.cal_per100!=null?db.cal_per100:db.cal)||0
    fat+=g*fv/100
    water+=g*wv/100
    freeWaterG+=g*fw/100
    flourEqG+=g*fe/100
    sugar+=g*sv/100
    protein+=g*pv/100
    carbs+=g*cv/100
    cal+=g*kv/100
    if(t==='salt')saltG+=g
  })
  if(!total)return null
  const bh=flourEqG>0?Math.round(freeWaterG/flourEqG*1000)/10:0
  const hydration=Math.round(water/total*1000)/10
  return{
    fat:Math.round(fat*10)/10,
    water:Math.round(water*10)/10,
    sugar:Math.round(sugar*10)/10,
    protein:Math.round(protein/total*1000)/10,
    carbs:Math.round(carbs/total*1000)/10,
    cal:Math.round(cal),
    total:Math.round(total),
    hydration,
    bakersHydration:bh,
    fatP:Math.round(fat/total*1000)/10,
    sugarP:Math.round(sugar/total*1000)/10,
    saltP:Math.round(saltG/total*1000)/10,
    saltG:Math.round(saltG*10)/10,
    flourEqG:Math.round(flourEqG),
    freeWaterG:Math.round(freeWaterG)
  }
}
function parseTabs(notesPad){
  if(!notesPad)return[{id:uid(),name:'General',content:''}]
  try{const p=JSON.parse(notesPad);if(Array.isArray(p)&&p.length>0)return p}catch(_){}
  return[{id:uid(),name:'General',content:notesPad}]
}

const serializeTabs = tabs => JSON.stringify(tabs)

/* ── I+D data helpers ───────────────────────────────────────────────────────────── */

function parseIdData(raw){
  if(!raw)return{sensory:{},versions:[],goal:'',nutritionOverride:null}
  try{return JSON.parse(raw)}catch{return{sensory:{},versions:[],goal:'',nutritionOverride:null}}
}

const serializeIdData = d => JSON.stringify(d)

/* ── Strip binary before API ────────────────────────────────────────────────────────── */

function stripForApi(recipe){
  if(!recipe)return recipe
  // eslint-disable-next-line no-unused-vars
  const{thumbnail,source_photos,media_library,...rest}=recipe
  return rest
}

/* ── Voice Input with baking correction ──────────────────────────────────── */

const BAKING_CORRECTIONS = {
  'krave':'krapfen','crave':'krapfen','grave':'krapfen',
  'macaroon':'macaron','macarons':'macaron',
  'brigadeiro':'brigadeiro','briggadeiro':'brigadeiro',
  'panettoni':'panettone','panetoni':'panettone',
  'focaccia':'focaccia','fokacia':'focaccia',
  'brioche':'brioche','brioce':'brioche','briosh':'brioche',
  'croissant':'croissant','cruasson':'croissant','cruasán':'croissant',
  'baguette':'baguette','baguet':'baguette',
  'choux':'choux','chu':'choux',
  'feuilletage':'feuilletage','feiyetage':'feuilletage',
  'laminage':'laminage',
  'autolyse':'autolyse','autolisis':'autolyse',
  'levain':'levain','levén':'levain',
  'poolish':'poolish','pulich':'poolish',
  'biga':'biga','viga':'biga',
  'sourdough':'sourdough','sour dough':'sourdough',
  'lievito naturale':'lievito naturale','levito natural':'lievito naturale',
  'maillard':'maillard','mayar':'maillard',
  'tangzhong':'tangzhong','tanjon':'tangzhong',
  'viennoiserie':'viennoiserie','vienoiserie':'viennoiserie',
  'pâte feuilletée':'pâte feuilletée',
  'crème pâtissière':'crème pâtissière','creme patissiere':'crème pâtissière',
  'ganache':'ganache','ganatche':'ganache',
  'praline':'praliné','pralinee':'praliné',
  'couverture':'couverture','covercure':'couverture',
}
function correctBakingTerms(text){
  let r=text
  Object.entries(BAKING_CORRECTIONS).forEach(([wrong,right])=>{
    const rx=new RegExp('\\b'+wrong+'\\b','gi')
    r=r.replace(rx,right)
  })
  return r
}
function useVoiceInput(onTranscript,smartCorrect=true){
  const[recording,setRecording]=useState(false)
  const ref=useRef(null);const acc=useRef('')
  function start(){
    const SR=window.SpeechRecognition||window.webkitSpeechRecognition
    if(!SR){alert('Voice input requires Chrome on desktop.');return}
    const r=new SR();r.lang=navigator.language||'en-US';r.continuous=true;r.interimResults=false
    r.onresult=e=>{const t=Array.from(e.results).slice(e.resultIndex).filter(x=>x.isFinal).map(x=>x[0].transcript).join(' ');if(t)acc.current+=(acc.current?' ':'')+t}
    r.onend=()=>{setRecording(false);if(acc.current.trim()){const final=smartCorrect?correctBakingTerms(acc.current.trim()):acc.current.trim();onTranscript(final);acc.current=''}};r.onerror=()=>setRecording(false)
    r.start();ref.current=r;acc.current='';setRecording(true)
  }
  function stop(){ref.current?.stop()}
  return{recording,start,stop}
}

/* ── Image Compression ──────────────────────────────────────────────────────────── */

function compressImage(file,maxW=1024,quality=0.7){
  return new Promise((res,rej)=>{
    const img=new Image(),url=URL.createObjectURL(file)
    img.onload=()=>{const scale=Math.min(1,maxW/img.width);const cv=document.createElement('canvas');cv.width=Math.round(img.width*scale);cv.height=Math.round(img.height*scale);cv.getContext('2d').drawImage(img,0,0,cv.width,cv.height);URL.revokeObjectURL(url);cv.toBlob(blob=>{if(!blob){rej(new Error('Failed'));return};const rd=new FileReader();rd.onload=()=>res({media_type:'image/jpeg',data:rd.result.split(',')[1],url:cv.toDataURL('image/jpeg',quality)});rd.onerror=rej;rd.readAsDataURL(blob)},'image/jpeg',quality)};img.onerror=()=>{URL.revokeObjectURL(url);rej(new Error('Load failed'))};img.src=url
  })
}
function compressThumbnail(file){
  return new Promise((res,rej)=>{
    const img=new Image(),url=URL.createObjectURL(file)
    img.onload=()=>{const s=Math.min(1,400/Math.max(img.width,img.height));const cv=document.createElement('canvas');cv.width=Math.round(img.width*s);cv.height=Math.round(img.height*s);cv.getContext('2d').drawImage(img,0,0,cv.width,cv.height);URL.revokeObjectURL(url);cv.toBlob(blob=>{if(!blob){rej(new Error('Failed'));return};const rd=new FileReader();rd.onload=()=>res(rd.result);rd.onerror=rej;rd.readAsDataURL(blob)},'image/jpeg',0.45)};img.onerror=()=>{URL.revokeObjectURL(url);rej(new Error('Load failed'))};img.src=url
  })
}
function loadImage(src){return new Promise(res=>{const img=new Image();img.onload=()=>res(img);img.onerror=()=>res(null);img.src=src})}

/* ── Media Library helpers ──────────────────────────────────────────────────────────── */

const MAX_AUDIO_BYTES = 2*60*128*1024/8  // ~2 min @ 128kbps = ~1.8MB
const MAX_VIDEO_BYTES = 3*60*1.5*1024*1024 // ~3 min @ 1.5MB/s = ~270MB → we cap at 30MB
const MAX_VIDEO_CAP   = 30*1024*1024
function readFileAsBase64(file){
  return new Promise((res,rej)=>{
    const rd=new FileReader();rd.onload=()=>res(rd.result);rd.onerror=rej;rd.readAsDataURL(file)
  })
}
function parseMediaLibrary(raw){
  if(!raw)return[]
  try{return JSON.parse(raw)}catch{return[]}
}

/* ── Edge Function calls ──────────────────────────────────────────────────────────── */

async function invoke(body){const{data,error}=await supabase.functions.invoke('extract-recipe',{body});if(error)throw new Error(error.message);if(data?.error)throw new Error(data.error);return data}

const extractWithClaude=imgs=>invoke({images:imgs})
const structureText=text=>invoke({type:'structure',text})
const translateRecipe=(recipe,lang)=>invoke({type:'translate',recipe:stripForApi(recipe),targetLang:lang})
const askAssistant=(msgs,recipe,lang)=>invoke({type:'assistant',messages:msgs.map(m=>({role:m.role,content:m.content})),recipe:stripForApi(recipe),language:lang})
const askAppAssistant=(msgs,recipes)=>invoke({type:'app_assistant',messages:msgs.map(m=>({role:m.role,content:m.content})),recipes:recipes.map(r=>({id:r.id,title:r.title,category:r.category,time:r.time,servings:r.servings}))})
const aiSuggestNotes=(recipe,currentNotes)=>invoke({type:'ai_suggest_notes',recipe:stripForApi(recipe),currentNotes})

const LANGS=['English','Spanish','French','Italian','German','Portuguese','Japanese']
const EXPIRY_OPTS={'24 h':86400000,'1 week':604800000,'1 month':2592000000}

/* ── Supabase CRUD ──────────────────────────────────────────────────────────── */

function toDb(r){return{
  id:r.id,title:r.title,category:r.category,time_estimate:r.time,
  servings:r.servings,notes:r.notes,source:r.source,
  ingredients:r.ingredients||[],steps:r.steps||[],
  notes_pad:r.notes_pad||'',thumbnail:r.thumbnail||'',
  source_photos:r.source_photos||[],
  id_data:r.id_data||'',
  media_library:r.media_library||'',
  fixed_lang:r.fixed_lang||null,
  copied_from:r.copied_from||null,
}}
function fromDb(r){return{...r,time:r.time_estimate,notes_pad:r.notes_pad||'',thumbnail:r.thumbnail||'',source_photos:r.source_photos||[],id_data:r.id_data||'',media_library:r.media_library||'',fixed_lang:r.fixed_lang||null,copied_from:r.copied_from||null}}
async function dbLoad(){const{data,error}=await supabase.from('recipes').select('*').order('created_at',{ascending:false});if(error)throw error;return(data||[]).map(fromDb)}
async function dbInsert(r){const p={...toDb(r)};delete p.id;const{data,error}=await supabase.from('recipes').insert([p]).select().single();if(error)throw error;return fromDb(data)}
async function dbUpdate(r){const{data,error}=await supabase.from('recipes').update(toDb(r)).eq('id',r.id).select().single();if(error)throw error;return fromDb(data)}
async function dbDelete(id){const{error}=await supabase.from('recipes').delete().eq('id',id);if(error)throw error}

/* ── PDF Export (fixed: thumbnail always from original recipe) ── */

function exportPDF(recipe,pctOpts=null,exportNotes=false,originalThumbnail=null){
  const doc=new jsPDF({unit:'mm',format:'a4'})
  const M=18,PW=210,CW=PW-M*2;let y=0
  function ck(n=12){if(y+n>279){doc.addPage();y=M}}
  doc.setFillColor(31,58,77);doc.rect(0,0,PW,4,'F');y=14
  // Use originalThumbnail if provided (fixes translation bug)
  const thumb=originalThumbnail||recipe.thumbnail
  if(thumb){try{doc.addImage(thumb,'JPEG',PW-M-28,8,28,28,'','FAST')}catch(_){}}
  const titleW=thumb?CW-34:CW
  doc.setFont('helvetica','bold');doc.setFontSize(22);doc.setTextColor(31,58,77)
  const tl=doc.splitTextToSize(recipe.title||'Recipe',titleW);doc.text(tl,M,y);y+=tl.length*9
  doc.setDrawColor(188,108,44);doc.setLineWidth(1.5);doc.line(M,y,M+28,y);y+=7
  const meta=[recipe.category&&`Category: ${recipe.category}`,recipe.time&&`Time: ${recipe.time}`,recipe.servings&&`Yield: ${recipe.servings}`].filter(Boolean)
  if(meta.length){doc.setFont('helvetica','normal');doc.setFontSize(9);doc.setTextColor(110,100,92);doc.text(meta.join('   ·   '),M,y);y+=9}
  if(thumb)y=Math.max(y,42)
  if(pctOpts?.appliedScaleLabel){doc.setFillColor(234,242,238);doc.rect(M,y,CW,6,'F');doc.setFont('helvetica','normal');doc.setFontSize(8.5);doc.setTextColor(45,106,79);doc.text(`⚖ Scaled — ${pctOpts.appliedScaleLabel}`,M+2,y+4.5);y+=8}
  const sections=parseSections(recipe.ingredients||[])
  y+=3;doc.setFont('helvetica','bold');doc.setFontSize(8);doc.setTextColor(31,58,77);doc.text('INGREDIENTS',M,y);y+=6
  sections.forEach(sec=>{
    if(sec.name){ck(12);y+=2;doc.setFont('helvetica','bolditalic');doc.setFontSize(9.5);doc.setTextColor(188,108,44);doc.text(sec.name,M,y);y+=6}
    const pctData=pctOpts?.showPct?calcPct(sec.items,pctOpts.pctMode,pctOpts.pctBase):null
    sec.items.forEach((ing,idx)=>{
      ck(7);const mm=ing.match(/^([\d.,]+\s*[^\s]+)\s{2,}(.+)$/)||ing.match(/^([\d.,]+\s*[a-zA-Z%]+)\s+(.+)$/)
      doc.setFont('helvetica','normal');doc.setFontSize(10);doc.setTextColor(34,28,24)
      if(mm){doc.setFont('courier','normal');doc.text(mm[1].trim(),M,y);doc.setFont('helvetica','normal');const ls=doc.splitTextToSize(mm[2].trim(),CW-38);doc.text(ls,M+36,y);if(pctData&&pctData[idx].pct!==null){doc.setFont('courier','normal');doc.setFontSize(9);doc.setTextColor(pctData[idx].isBase?188:110,pctData[idx].isBase?108:100,pctData[idx].isBase?44:92);doc.text(pctData[idx].pct.toFixed(1)+'%',PW-M,y,{align:'right'});doc.setTextColor(34,28,24);doc.setFontSize(10)};y+=ls.length*5+1}
      else{const ls=doc.splitTextToSize(`· ${ing}`,CW);doc.text(ls,M,y);y+=ls.length*5+1}
    })
    const sg=sec.items.reduce((s,i)=>{const p=parseIng(i);return s+toGrams(p.qty,p.unit)},0)
    if(sec.name&&sg>0){doc.setFont('courier','normal');doc.setFontSize(8.5);doc.setTextColor(110,100,92);doc.text(`Subtotal: ${sg.toFixed(0)} g`,PW-M,y,{align:'right'});y+=6}
  })
  const tg=getTotalGrams(recipe.ingredients||[])
  if(tg>0){doc.setFont('helvetica','bold');doc.setFontSize(9);doc.setTextColor(31,58,77);doc.text(`Total: ${tg.toFixed(0)} g`,PW-M,y,{align:'right'});y+=8}
  if(recipe.steps?.length){ck(12);y+=3;doc.setFont('helvetica','bold');doc.setFontSize(8);doc.setTextColor(31,58,77);doc.text('METHOD',M,y);y+=6;recipe.steps.forEach((step,i)=>{ck(10);const ls=doc.splitTextToSize(step,CW-14);doc.setFont('helvetica','bold');doc.setFontSize(10);doc.setTextColor(188,108,44);doc.text(String(i+1).padStart(2,'0')+'.',M,y);doc.setFont('helvetica','normal');doc.setTextColor(34,28,24);doc.text(ls,M+14,y);y+=ls.length*5.5+3})}
  if(recipe.notes){ck(18);y+=3;const nl=doc.splitTextToSize(recipe.notes,CW-9);const bh=nl.length*5.5+10;doc.setFillColor(251,239,225);doc.setDrawColor(188,108,44);doc.setLineWidth(.2);doc.rect(M,y,CW,bh,'FD');doc.setFillColor(188,108,44);doc.rect(M,y,2.5,bh,'F');doc.setFont('helvetica','italic');doc.setFontSize(9.5);doc.setTextColor(110,100,92);doc.text(nl,M+5,y+7);y+=bh}
  if(exportNotes){
    const tabs=parseTabs(recipe.notes_pad).filter(t=>t.content.trim())
    if(tabs.length){ck(16);y+=5;doc.setFont('helvetica','bold');doc.setFontSize(8);doc.setTextColor(31,58,77);doc.text('NOTES',M,y);y+=6;tabs.forEach(tab=>{if(tab.name!=='General'){ck(8);doc.setFont('helvetica','bolditalic');doc.setFontSize(9);doc.setTextColor(91,58,140);doc.text(tab.name,M,y);y+=5};const lines=tab.content.split('\n');lines.forEach(line=>{ck(6);const ls=doc.splitTextToSize(line||' ',CW);doc.setFont('helvetica','normal');doc.setFontSize(9.5);doc.setTextColor(110,100,92);doc.text(ls,M,y);y+=ls.length*5});y+=3})}
  }
  const total=doc.internal.getNumberOfPages()
  for(let p=1;p<=total;p++){doc.setPage(p);doc.setFont('helvetica','normal');doc.setFontSize(7);doc.setTextColor(180,170,160);doc.text('Quaderno AI+D',M,292);doc.text(`${new Date().toLocaleDateString()}  ·  ${p}/${total}`,PW-M,292,{align:'right'});doc.setFillColor(31,58,77);doc.rect(0,294,PW,2,'F')}
  doc.save((recipe.title||'recipe').replace(/[^a-z0-9]+/gi,'-').toLowerCase()+'.pdf')
}

/* ── Image Export (fixed: thumbnail from original) ─────────────────────── */

async function exportImage(recipe,pctOpts=null,exportNotes=false,originalThumbnail=null){
  const thumb=originalThumbnail||recipe.thumbnail
  const thumbImg=thumb?await loadImage(thumb):null
  const W=1200,M=68,CW=W-M*2,DPR=2,TH=7000
  const cv=document.createElement('canvas');cv.width=W*DPR;cv.height=TH*DPR
  const ctx=cv.getContext('2d');ctx.scale(DPR,DPR)
  ctx.fillStyle='#FAF7F0';ctx.fillRect(0,0,W,TH);let y=0
  const gl=(text,maxW)=>{const wds=String(text||'').split(' ');const ls=[];let l='';for(const w of wds){const t=l?l+' '+w:w;if(ctx.measureText(t).width>maxW&&l){ls.push(l);l=w}else l=t};if(l)ls.push(l);return ls}
  const dt=(text,x,yy,maxW,lh)=>{const ls=gl(text,maxW);ls.forEach((l,i)=>ctx.fillText(l,x,yy+i*lh));return ls.length*lh}
  function rr(x,yy,w,h,r){ctx.beginPath();ctx.moveTo(x+r,yy);ctx.arcTo(x+w,yy,x+w,yy+h,r);ctx.arcTo(x+w,yy+h,x,yy+h,r);ctx.arcTo(x,yy+h,x,yy,r);ctx.arcTo(x,yy,x+w,yy,r);ctx.closePath()}
  ctx.fillStyle='#1F3A4D';ctx.fillRect(0,0,W,10);y=62
  if(thumbImg){const SZ=110,TX=W-M-SZ,TY=16;ctx.save();rr(TX,TY,SZ,SZ,8);ctx.clip();ctx.drawImage(thumbImg,TX,TY,SZ,SZ);ctx.restore();ctx.strokeStyle='rgba(230,222,207,.8)';ctx.lineWidth=1;rr(TX,TY,SZ,SZ,8);ctx.stroke()}
  ctx.font='bold 44px Georgia,serif';ctx.fillStyle='#1F3A4D'
  const titleMaxW=thumbImg?CW-130:CW
  y+=dt(recipe.title||'Recipe',M,y,titleMaxW,56)
  ctx.strokeStyle='#BC6C2C';ctx.lineWidth=3;ctx.beginPath();ctx.moveTo(M,y+5);ctx.lineTo(M+110,y+5);ctx.stroke();y+=22
  if(thumbImg)y=Math.max(y,140)
  const mp=[recipe.category,recipe.time&&`⏱ ${recipe.time}`,recipe.servings&&`⚖ ${recipe.servings}`].filter(Boolean)
  if(mp.length){ctx.font='17px -apple-system,sans-serif';ctx.fillStyle='#6E645C';ctx.fillText(mp.join('   ·   '),M,y);y+=34}
  if(pctOpts?.appliedScaleLabel){ctx.font='bold 14px ui-monospace,monospace';ctx.fillStyle='#2D6A4F';ctx.fillText('⚖ '+pctOpts.appliedScaleLabel,M,y);y+=24}
  y+=10;ctx.strokeStyle='#E6DECF';ctx.lineWidth=1;ctx.beginPath();ctx.moveTo(M,y);ctx.lineTo(W-M,y);ctx.stroke();y+=28
  const secLbl=(label,off)=>{ctx.font='bold 11px ui-monospace,monospace';ctx.fillStyle='#1F3A4D';ctx.fillText(label,M,y);ctx.strokeStyle='#1F3A4D';ctx.lineWidth=.5;ctx.beginPath();ctx.moveTo(M+off,y-3);ctx.lineTo(W-M,y-3);ctx.stroke();y+=24}
  const sections=parseSections(recipe.ingredients||[])
  secLbl('INGREDIENTS',112)
  sections.forEach(sec=>{
    if(sec.name){ctx.font='bold italic 17px Georgia,serif';ctx.fillStyle='#BC6C2C';ctx.fillText(sec.name,M,y);y+=28}
    const pctData=pctOpts?.showPct?calcPct(sec.items,pctOpts.pctMode,pctOpts.pctBase):null
    sec.items.forEach((ing,idx)=>{
      const mm=ing.match(/^([\d.,]+\s*[^\s]+)\s{2,}(.+)$/)||ing.match(/^([\d.,]+\s*[a-zA-Z%]+)\s+(.+)$/)
      const pct=pctData?pctData[idx]:null;const nW=pct?.pct!=null?CW-320:CW-186
      if(mm){ctx.font='16px ui-monospace,monospace';ctx.fillStyle='#6E645C';ctx.fillText(mm[1].trim(),M,y);ctx.font='17px -apple-system,sans-serif';ctx.fillStyle='#221C18';const used=dt(mm[2].trim(),M+186,y,nW,27);if(pct?.pct!=null){ctx.font='bold 15px ui-monospace,monospace';ctx.fillStyle=pct.isBase?'#BC6C2C':'#6E645C';ctx.textAlign='right';ctx.fillText(pct.pct.toFixed(1)+'%',W-M,y);ctx.textAlign='left'};y+=Math.max(27,used)+3}
      else{ctx.font='17px -apple-system,sans-serif';ctx.fillStyle='#221C18';y+=dt(`· ${ing}`,M+10,y,CW-10,27)+3}
    })
    const sg=sec.items.reduce((s,i)=>{const p=parseIng(i);return s+toGrams(p.qty,p.unit)},0)
    if(sec.name&&sg>0){ctx.font='12px ui-monospace,monospace';ctx.fillStyle='#BC6C2C';ctx.textAlign='right';ctx.fillText(`Subtotal: ${sg.toFixed(0)} g`,W-M,y);ctx.textAlign='left';y+=18}
  })
  const tg=getTotalGrams(recipe.ingredients||[])
  if(tg>0){ctx.font='bold 13px ui-monospace,monospace';ctx.fillStyle='#1F3A4D';ctx.textAlign='right';ctx.fillText(`Total: ${tg.toFixed(0)} g`,W-M,y);ctx.textAlign='left';y+=26}
  if(recipe.steps?.length){secLbl('METHOD',76);recipe.steps.forEach((step,i)=>{ctx.font='bold 17px -apple-system,sans-serif';ctx.fillStyle='#BC6C2C';ctx.fillText(String(i+1).padStart(2,'0')+'.',M,y);ctx.font='17px -apple-system,sans-serif';ctx.fillStyle='#221C18';y+=Math.max(27,dt(step,M+46,y,CW-46,27))+7})}
  if(recipe.notes){ctx.font='italic 15px -apple-system,sans-serif';const nl=gl(recipe.notes,CW-44);const bh=nl.length*27+36;rr(M,y,CW,bh,8);ctx.fillStyle='#FBEFE1';ctx.fill();ctx.fillStyle='#BC6C2C';ctx.fillRect(M,y,4,bh);ctx.fillStyle='#6E645C';nl.forEach((l,i)=>ctx.fillText(l,M+16,y+24+i*27));y+=bh+20}
  if(exportNotes){
    const tabs=parseTabs(recipe.notes_pad).filter(t=>t.content.trim())
    if(tabs.length){y+=14;ctx.font='bold 11px ui-monospace,monospace';ctx.fillStyle='#1F3A4D';ctx.fillText('NOTES',M,y);ctx.strokeStyle='#1F3A4D';ctx.lineWidth=.5;ctx.beginPath();ctx.moveTo(M+56,y-3);ctx.lineTo(W-M,y-3);ctx.stroke();y+=22;tabs.forEach(tab=>{if(tab.name!=='General'){ctx.font='bold italic 16px Georgia,serif';ctx.fillStyle='#5B3A8C';ctx.fillText(tab.name,M,y);y+=24};ctx.font='italic 15px -apple-system,sans-serif';ctx.fillStyle='#6E645C';tab.content.split('\n').forEach(line=>{if(line.trim()){y+=dt(line,M+10,y,CW-10,26)}else y+=13})})}
  }
  y+=12;ctx.strokeStyle='#E6DECF';ctx.lineWidth=1;ctx.beginPath();ctx.moveTo(M,y);ctx.lineTo(W-M,y);ctx.stroke();y+=20
  ctx.font='12px ui-monospace,monospace';ctx.fillStyle='#B0A89F';ctx.fillText('Quaderno AI+D',M,y);ctx.textAlign='right';ctx.fillText(new Date().toLocaleDateString('en-US',{year:'numeric',month:'long',day:'numeric'}),W-M,y);ctx.textAlign='left';y+=18;ctx.fillStyle='#1F3A4D';ctx.fillRect(0,y,W,10);y+=10
  const out=document.createElement('canvas');out.width=W*DPR;out.height=y*DPR;out.getContext('2d').drawImage(cv,0,0,W*DPR,y*DPR,0,0,W*DPR,y*DPR)
  const a=document.createElement('a');a.href=out.toDataURL('image/png');a.download=(recipe.title||'recipe').replace(/[^a-z0-9]+/gi,'-').toLowerCase()+'.png';document.body.appendChild(a);a.click();document.body.removeChild(a)
}

/* ── XLS Export ──────────────────────────────────────────────────────────── */

function exportXLS(recipe,pctOpts=null){
  const wb=XLSX.utils.book_new()
  const rows=[[`QUADERNO AI+D — `+(recipe.title||'Recipe')],[],['Category',recipe.category||''],['Time',recipe.time||''],['Yield',recipe.servings||''],[],[pctOpts?.showPct?"INGREDIENTS (Baker's %)":"INGREDIENTS"]]
  const sections=parseSections(recipe.ingredients||[]);const IROW=rows.length+1;const flourRows=[];let curRow=IROW
  sections.forEach(sec=>{if(sec.name){rows.push(['── '+sec.name+' ──','','','']);curRow++};sec.items.forEach(ing=>{const p=parseIng(ing);const g=toGrams(p.qty,p.unit);rows.push([p.name,p.qty||'',p.unit||'',g>0?g:'']);if(isFlour(p.name))flourRows.push(curRow);curRow++});const sg=sec.items.reduce((s,i)=>{const p=parseIng(i);return s+toGrams(p.qty,p.unit)},0);if(sec.name&&sg>0){rows.push(['Subtotal: '+sec.name,'','',sg]);curRow++}})
  const tg=getTotalGrams(recipe.ingredients||[]);if(tg>0)rows.push(['TOTAL','','',tg]);rows.push([],['METHOD']);(recipe.steps||[]).forEach((s,i)=>rows.push([(i+1)+'.',s]))
  if(recipe.notes){rows.push([],["Baker's Notes"]);rows.push([recipe.notes])}
  const ws1=XLSX.utils.aoa_to_sheet(rows)
  ws1['!cols']=[{wch:32},{wch:10},{wch:8},{wch:10},{wch:12}];XLSX.utils.book_append_sheet(wb,ws1,'Recipe')
  // Baker's Calc sheet
  const flatItems=sections.flatMap(s=>s.items);const calcRows=[["Baker's Calculator — "+(recipe.title||'')],['Edit grams in column B → percentages update automatically.'],[],['Ingredient','Grams',"Baker's %",'Note']]
  const calcFlourRows=[];flatItems.forEach((ing,i)=>{const p=parseIng(ing);const g=toGrams(p.qty,p.unit);if(isFlour(p.name))calcFlourRows.push(5+i);calcRows.push([p.name,g>0?g:0,'',isFlour(p.name)?'← flour base':''])})
  const ws2=XLSX.utils.aoa_to_sheet(calcRows)
  ws2['!cols']=[{wch:32},{wch:12},{wch:12},{wch:14}];XLSX.utils.book_append_sheet(wb,ws2,"Baker's Calc")
  // Macros sheet
  const m=calcMacros(recipe.ingredients)
  const macroRows=[['I+D — Macros & Parameters',''],['Recipe',recipe.title||''],['Total batch',m.total+'g'],[''],['MACRO PARAMETERS','Per 100g of batch'],['Total fat',m.fatP+'%'],['Total water / hydration',m.hydration+'%'],["Baker's hydration",m.bakersHydration+'%'],['Total sugar',m.sugarP+'%'],["Salt (baker's %)",m.saltP+'%'],[''],['NUTRITIONAL (estimate)','Per batch'],['Fat',m.fat+'g'],['Water',m.water+'g'],['Sugar',m.sugar+'g'],['Protein',m.protein+'g'],['Carbs',m.carbs+'g'],['Calories (kcal)',m.cal]]
  const ws3=XLSX.utils.aoa_to_sheet(macroRows)
  ws3['!cols']=[{wch:28},{wch:18}];XLSX.utils.book_append_sheet(wb,ws3,'I+D Macros')
  XLSX.writeFile(wb,(recipe.title||'recipe').replace(/[^a-z0-9]+/gi,'-').toLowerCase()+'.xlsx')
}

/* ── DraggableIngList ──────────────────────────────────────────────────────────── */

function DraggableIngList({lines,onChange}){
  const[dragIdx,setDragIdx]=useState(null)
  const[overIdx,setOverIdx]=useState(null)
  function move(from,to){if(from===to)return;const n=[...lines];const[item]=n.splice(from,1);n.splice(to,0,item);onChange(n);setDragIdx(null);setOverIdx(null)}
  const isSec=l=>/^##?\s+/.test(l)
  return(
    <div className="Q-drag-list">
      {lines.map((line,idx)=>(
        <div key={idx} className={`Q-drag-item${overIdx===idx?' over':''}${dragIdx===idx?' dragging':''}${isSec(line)?' is-section':''}`}
          draggable onDragStart={e=>{setDragIdx(idx);e.dataTransfer.effectAllowed='move'}}
          onDragOver={e=>{e.preventDefault();setOverIdx(idx)}} onDrop={()=>dragIdx!==null&&move(dragIdx,idx)}
          onDragEnd={()=>{setDragIdx(null);setOverIdx(null)}}>
          <span className="Q-drag-handle">⠇</span>
          <input className={`Q-drag-input${isSec(line)?' section':''}`}
            value={isSec(line)?line.replace(/^##?\s*/,''):line}
            onChange={e=>{const n=[...lines];n[idx]=isSec(line)?'## '+e.target.value:e.target.value;onChange(n)}}
            placeholder={isSec(line)?'Section name':'500 g  ingredient name'}/>
          <button className="Q-drag-rm" onClick={()=>onChange(lines.filter((_,i)=>i!==idx))}>×</button>
        </div>
      ))}
      <div className="Q-drag-footer">
        <button onClick={()=>onChange([...lines,''])} style={{fontSize:11.5,background:'none',border:'1px solid var(--rule)',borderRadius:5,padding:'4px 10px',cursor:'pointer',color:'var(--muted)'}}>+ Ingredient</button>
        <button onClick={()=>onChange([...lines,'## '])} style={{fontSize:11.5,background:'none',border:'1px solid var(--amber)',borderRadius:5,padding:'4px 10px',cursor:'pointer',color:'var(--amber)'}}>+ Section</button>
      </div>
    </div>
  )
}

/* ── Media Library Panel ──────────────────────────────────────────────────────────── */

function MediaLibraryPanel({recipeId,mediaRaw,onSave}){
  const[items,setItems]=useState(()=>parseMediaLibrary(mediaRaw))
  const[uploading,setUploading]=useState(false)
  const[err,setErr]=useState('')
  const[player,setPlayer]=useState(null) // {type,src,name}
  const fileRef=useRef(null)
  useEffect(()=>{setItems(parseMediaLibrary(mediaRaw))},[recipeId])
  async function handleFiles(files){
    setErr('');setUploading(true)
    try{
      const newItems=[]
      for(const f of Array.from(files)){
        const isImg=f.type.startsWith('image/')
        const isAudio=f.type.startsWith('audio/')
        const isVideo=f.type.startsWith('video/')
        if(!isImg&&!isAudio&&!isVideo){setErr('Only images, audio, and video supported.');continue}
        if(isAudio&&f.size>MAX_AUDIO_BYTES){setErr(`Audio max ~2 min (file too large: ${(f.size/1024/1024).toFixed(1)}MB)`);continue}
        if(isVideo&&f.size>MAX_VIDEO_CAP){setErr(`Video max 30 MB. Please compress before uploading.`);continue}
        if(isImg){
          const compressed=await compressImage(f,1200,0.8)
          newItems.push({id:uid(),type:'image',src:compressed.url,name:f.name,date:new Date().toISOString()})
        } else {
          const b64=await readFileAsBase64(f)
          newItems.push({id:uid(),type:isAudio?'audio':'video',src:b64,name:f.name,date:new Date().toISOString(),mime:f.type})
        }
      }
      const updated=[...items,...newItems]
      setItems(updated)
      await onSave(JSON.stringify(updated))
    }catch(e){setErr('Upload failed: '+e.message)}
    finally{setUploading(false)}
  }
  async function remove(id){
    const updated=items.filter(i=>i.id!==id)
    setItems(updated)
    await onSave(JSON.stringify(updated))
  }
  return(
    <div>
      <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:10}}>
        <span style={{fontFamily:'var(--mono)',fontSize:9.5,textTransform:'uppercase',letterSpacing:'.18em',color:' var(--navy)'}}>Media Library — {items.length} file{items.length!==1?'s':''}</span>
        <label className="Q-media-upload-btn">
          {uploading?'Uploading…':'＋ Add media'}
          <input ref={fileRef} type="file" multiple accept="image/*,audio/*,video/*" style={{display:'none'}} disabled={uploading} onChange={e=>handleFiles(e.target.files)}/>
        </label>
      </div>
      {err&&<div className="Q-err" style={{marginBottom:8}}>{err}</div>}
      <div style={{fontSize:11,color:'var(--muted)',marginBottom:10,lineHeight:1.5}}>
        Photos · Audio ≤ 2 min · Video ≤ 30 MB (3 min approx)
      </div>
      {items.length===0&&<div style={{textAlign:'center',color:'var(--muted)',padding:'32px 0',fontSize:13}}>No media yet. Add photos, audio recordings, or videos.</div>}
      <div className="Q-media-grid">
        {items.map(item=>(
          <div key={item.id} className="Q-media-card" onClick={()=>item.type!=='image'&&setPlayer(item)}>
            {item.type==='image'&&<img src={item.src} className="Q-media-thumb" alt={item.name} onClick={()=>setPlayer(item)}/>}
            {item.type==='audio'&&<div className="Q-media-audio"><span style={{fontSize:24}}>🎙</span><span style={{fontSize:10,color:'var(--muted)'}}>Audio</span></div>}
            {item.type==='video'&&<div className="Q-media-video"><span style={{fontSize:24}}>🎬</span><span style={{fontSize:10}}>Video</span></div>}
            <div className="Q-media-label">{item.name.length>16?item.name.slice(0,13)+'…':item.name}</div>
            <button className="Q-media-rm" onClick={e=>{e.stopPropagation();remove(item.id)}}>×</button>
          </div>
        ))}
      </div>
      {player&&(
        <div className="Q-media-player-wrap" onClick={()=>setPlayer(null)}>
          <div onClick={e=>e.stopPropagation()} style={{display:'flex',flexDirection:'column',alignItems:'center',gap:12}}>
            {player.type==='image'&&<img src={player.src} style={{maxWidth:'90vw',maxHeight:'80vh',borderRadius:8}} alt=""/>}
            {player.type==='audio'&&<audio controls autoPlay src={player.src}/>}
            {player.type==='video'&&<video controls autoPlay style={{maxWidth:'90vw',maxHeight:'80vh'}} src={player.src}/>}
            <div style={{fontFamily:'var(--mono)',fontSize:11,color:'rgba(255,255,255,.7)'}}>{player.name}</div>
            <button className="btn ghost xs" onClick={()=>setPlayer(null)}>✕ Close</button>
          </div>
        </div>
      )}
    </div>
  )
}

/* ── Notes Panel (tabs + media) ──────────────────────────────────────────────────────────── */

function NotesPanel({recipe,onSave,onSaveMedia,onAddNote}){
  const[tabs,setTabs]=useState(()=>parseTabs(recipe.notes_pad))
  const[activeIdx,setActiveIdx]=useState(0)
  const[activeSection,setActiveSection]=useState('notes') // 'notes' | 'media'
  const[saving,setSaving]=useState(false)
  const[aiLoading,setAiLoading]=useState(false)
  const timer=useRef(null)
  useEffect(()=>{setTabs(parseTabs(recipe.notes_pad));setActiveIdx(0)},[recipe.id])
  useEffect(()=>{if(onAddNote)onAddNote.current=appendNote},[tabs,activeIdx])
  function appendNote(content){
    setTabs(prev=>{const next=prev.map((t,i)=>i===activeIdx?{...t,content:t.content+(t.content?'\n\n':'')+content}:t);save(next);return next})
  }
  function save(t){clearTimeout(timer.current);timer.current=setTimeout(async()=>{setSaving(true);try{await onSave(serializeTabs(t))}finally{setSaving(false)}},1200)}
  const voice=useVoiceInput(transcript=>{
    const entry=`[${ts()}]\n${transcript}`
    setTabs(prev=>{const next=prev.map((t,i)=>i===activeIdx?{...t,content:t.content+(t.content?'\n\n':'')+entry}:t);save(next);return next})
  },true)
  function updateContent(val){setTabs(prev=>{const next=prev.map((t,i)=>i===activeIdx?{...t,content:val}:t);save(next);return next})}
  function addTab(){const newTab={id:uid(),name:'New Tab',content:''};const next=[...tabs,newTab];setTabs(next);setActiveIdx(next.length-1);save(next)}
  function renameTab(idx){const name=prompt('Tab name:',tabs[idx].name);if(name&&name.trim()){setTabs(prev=>{const next=prev.map((t,i)=>i===idx?{...t,name:name.trim()}:t);save(next);return next})}}
  function removeTab(idx){if(tabs.length===1){alert('Cannot remove the only tab.');return};if(!confirm(`Remove tab "${tabs[idx].name}"?`))return;const next=tabs.filter((_,i)=>i!==idx);setTabs(next);setActiveIdx(Math.min(activeIdx,next.length-1));save(next)}
  async function aiSuggest(){setAiLoading(true);try{const r=await aiSuggestNotes(recipe,tabs[activeIdx]?.content||'');const t=r?.text||'';appendNote('✨ AI Suggestions:\n'+t)}catch(e){alert('AI suggest failed: '+e.message)}finally{setAiLoading(false)}}
  const activeTab=tabs[activeIdx]||tabs[0]
  return(
    <div className="Q-notes-panel">
      <div style={{display:'flex',gap:0,marginBottom:12,borderBottom:'2px solid var(--rule)'}}>
        {[['notes','📝 Notes'],['media','🖼 Media']].map(([k,l])=>(
          <button key={k} onClick={()=>setActiveSection(k)} style={{fontFamily:'var(--sans)',fontSize:12.5,fontWeight:600,background:'none',border:'none',padding:'8px 14px',cursor:'pointer',color:activeSection===k?'var(--navy)':'var(--muted)',borderBottom:activeSection===k?'2px solid var(--amber)':'2px solid transparent',marginBottom:-2}}>
            {l}
          </button>
        ))}
      </div>
      {activeSection==='notes'&&<>
        <div className="Q-notes-tabs-row">
          {tabs.map((t,i)=>(
            <div key={t.id} style={{display:'flex',alignItems:'center'}}>
              <button className={`Q-notes-tab${i===activeIdx?' active':''}`} onClick={()=>setActiveIdx(i)}>{t.name}</button>
              {i===activeIdx&&<button onClick={()=>renameTab(i)} title="Rename" style={{background:'none',border:'none',cursor:'pointer',fontSize:11,color:'var(--muted)',padding:'0 3px',lineHeight:1,marginLeft:-2}}>✏</button>}
            </div>
          ))}
          <button className="Q-notes-tab-add" onClick={addTab} title="Add tab">＋</button>
        </div>
        <div className="Q-notes-toolbar">
          <span style={{fontFamily:'var(--mono)',fontSize:9.5,textTransform:'uppercase',letterSpacing:'.18em',color:'var(--navy)'}}>📝 {activeTab?.name}</span>
          <div style={{display:'flex',gap:6,alignItems:'center',flexWrap:'wrap'}}>
            {saving&&<span style={{fontSize:10,color:'var(--muted)',fontFamily:'var(--mono)'}}>saving…</span>}
            {voice.recording&&<span className="Q-recording-pill">🔴 Recording…</span>}
            {tabs.length>1&&<button className="btn danger xs" onClick={()=>removeTab(activeIdx)}>Remove tab</button>}
            <button className={`Q-voice-btn${voice.recording?' recording':''}`} onClick={voice.recording?voice.stop:voice.start} title="Voice note">{voice.recording?'⏹':'🎙'}</button>
            <button className="btn xs ai" onClick={aiSuggest} disabled={aiLoading}>{aiLoading?'…':'✨ AI'}</button>
          </div>
        </div>
        {voice.recording&&<div style={{fontSize:11,color:'var(--muted)',marginBottom:6}}>🔴 Recording… tap ⏹ to stop. Transcribed + timestamp added.</div>}
        <textarea className="Q-notes-textarea"
          value={activeTab?.content||''}
          onChange={e=>updateContent(e.target.value)}
          placeholder={`${activeTab?.name} for ${recipe.title}…\n\n🎙 Mic → transcribes with smart baking term correction + timestamp\n✨ AI → suggestions based on this recipe`}/>
      </>}
      {activeSection==='media'&&<MediaLibraryPanel recipeId={recipe.id} mediaRaw={recipe.media_library||''} onSave={onSaveMedia}/>}
    </div>
  )
}

/* ── I+D Panel ──────────────────────────────────────────────────────────── */

const SENSORY_ATTRS=['Appearance','Color','Aroma','Texture','Crumb structure','Crust','Flavor','Aftertaste','Overall score']
const SENSORY_LABELS={
  'Appearance':'Visual presentation','Color':'Color uniformity','Aroma':'Smell intensity & character',
  'Texture':'Mouthfeel & bite','Crumb structure':'Open, uniform, or tight crumb','Crust':'Thickness, color, crunch',
  'Flavor':'Taste balance & complexity','Aftertaste':'Persistence & pleasantness','Overall score':'Overall quality'
}
function IDPanel({recipe,onSave}){
  const[data,setData]=useState(()=>parseIdData(recipe.id_data))
  const[openSections,setOpenSections]=useState({params:true,sensory:false,timeline:false,nutrition:false})
  const[newVersion,setNewVersion]=useState({title:'',notes:'',isGoal:false})
  const[addingVersion,setAddingVersion]=useState(false)
  const[saving,setSaving]=useState(false)
  useEffect(()=>{setData(parseIdData(recipe.id_data))},[recipe.id])
  async function saveData(d){setSaving(true);try{await onSave(serializeIdData(d))}finally{setSaving(false)}}
  function toggleSection(k){setOpenSections(p=>({...p,[k]:!p[k]}))}
  function setSensory(attr,field,val){
    const next={...data,sensory:{...data.sensory,[attr]:{...(data.sensory[attr]||{}), [field]:val}}}
    setData(next);saveData(next)
  }
  function addVersion(){
    if(!newVersion.title.trim())return
    const v={id:uid(),date:new Date().toISOString().slice(0,10),title:newVersion.title.trim(),notes:newVersion.notes.trim(),isGoal:newVersion.isGoal}
    const next={...data,versions:[...(data.versions||[]),v]}
    setData(next);saveData(next);setNewVersion({title:'',notes:'',isGoal:false});setAddingVersion(false)
  }
  function removeVersion(id){
    const next={...data,versions:(data.versions||[]).filter(v=>v.id!==id)}
    setData(next);saveData(next)
  }
  function setGoal(g){const next={...data,goal:g};setData(next);saveData(next)}
  const macros=useMemo(()=>calcMacros(recipe.ingredients,data.macroCache||null),[recipe.ingredients,data.macroCache])
  const MAX_VALS={fat:40,water:80,sugar:40,protein:20,bakersHydration:100,saltP:3}
  const [analyzingMacros,setAnalyzingMacros]=React.useState(false)
  const [addingParam,setAddingParam]=React.useState(false)
  const [newParamLabel,setNewParamLabel]=React.useState('')
  const [analyzingCustom,setAnalyzingCustom]=React.useState(false)
  async function runAIMacroAnalysis(){
    if(!recipe.ingredients||!recipe.ingredients.length)return
    setAnalyzingMacros(true)
    try{
      const ingList=recipe.ingredients.filter(i=>!/^##?\s+/.test(i)).map(i=>parseIng(i)).filter(p=>p.qty!=null).map(p=>({name:p.name,grams:toGrams(p.qty,p.unit)||0}))
      const {data:_aiData,error:_aiErr}=await supabase.functions.invoke('extract-recipe',{body:{action:'analyze_macros',ingredients:ingList}});if(_aiErr)throw _aiErr
      const json=_aiData
      if(json.cache){
        const next={...data,macroCache:json.cache}
        setData(next);saveData(next)
      }
    }catch(e){console.error('AI macro analysis failed:',e)}
    setAnalyzingMacros(false)
  }
  async function addCustomParam(){
    if(!newParamLabel.trim())return
    setAnalyzingCustom(true)
    try{
      const ingList=recipe.ingredients.filter(i=>!/^##?\s+/.test(i)).map(i=>parseIng(i)).filter(p=>p.qty!=null).map(p=>({name:p.name,grams:toGrams(p.qty,p.unit)||0}))
      const {data:_cpData,error:_cpErr}=await supabase.functions.invoke('extract-recipe',{body:{action:'analyze_custom_param',param_label:newParamLabel.trim(),ingredients:ingList,existing_macros:macros}});if(_cpErr)throw _cpErr
      const json=_cpData
      const np={id:Date.now(),label:newParamLabel.trim(),value:json.value??0,unit:json.unit||'',explanation:json.explanation||''}
      const updated=[...(data.customParams||[]),np]
      const next={...data,customParams:updated}
      setData(next);saveData(next)
      setNewParamLabel('');setAddingParam(false)
    }catch(e){console.error('Custom param failed:',e)}
    setAnalyzingCustom(false)
  }
  function removeCustomParam(id){
    const updated=(data.customParams||[]).filter(p=>p.id!==id)
    const next={...data,customParams:updated}
    setData(next);saveData(next)
  }
  return(
    <div className="ID-panel">
      {saving&&<div style={{fontFamily:'var(--mono)',fontSize:10,color:'var(--id)',marginBottom:8}}>Saving…</div>}
      {/* ─ Parameters ─ */}
      <div className="ID-section">
        <div className="ID-section-header" onClick={()=>toggleSection('params')}>
          <h3>📊 Macro Parameters</h3>
          <span style={{fontSize:11,color:'var(--id)'}}>{openSections.params?'▲':'▼'}</span>
          <span style={{fontFamily:'var(--mono)',fontSize:9,color:'var(--muted)',marginLeft:8}}>{data.macroCache?'AI-powered':'auto-calculated'}</span>
        </div>
        {openSections.params&&<div className="ID-section-body">
          <div style={{display:'flex',gap:6,marginBottom:8,flexWrap:'wrap'}}>
            <button onClick={runAIMacroAnalysis} disabled={analyzingMacros} style={{fontSize:10,padding:'2px 8px',borderRadius:4,border:'1px solid var(--id)',background:'transparent',color:'var(--id)',cursor:'pointer'}}>
              {analyzingMacros?'Analyzing...':'\uD83E\uDD16 AI analyze'}
            </button>
            {data.macroCache&&<span style={{fontSize:9,color:'var(--muted)',alignSelf:'center'}}>AI cache active</span>}
          </div>
          <div className="ID-param-grid">
            {[{k:'fat',label:'Total fat',v:macros&&macros.fatP,unit:'%',max:MAX_VALS.fat,color:'#E07B39'},
              {k:'water',label:'Hydration (total)',v:macros&&macros.hydration,unit:'%',max:MAX_VALS.water,color:'#2196F3'},
              {k:'bakersHydration',label:"Baker's hydration",v:macros&&macros.bakersHydration,unit:'%',max:MAX_VALS.bakersHydration,color:'#1A6B6B'},
              {k:'sugar',label:'Total sugar',v:macros&&macros.sugarP,unit:'%',max:MAX_VALS.sugar,color:'#9C27B0'},
              {k:'saltP',label:"Salt (baker's %)",v:macros&&macros.saltP,unit:'%',max:MAX_VALS.saltP,color:'#607D8B'},
              {k:'protein',label:'Protein (est.)',v:macros&&macros.protein,unit:'g',max:null,color:'#2D6A4F'},
            ].map(p=>(
              <div className="ID-param-card" key={p.k}>
                <div className="ID-param-label">{p.label}</div>
                <div className="ID-param-value">{p.v}<span className="ID-param-unit">{p.unit}</span></div>
                {p.max&&<div className="ID-param-bar"><div className="ID-param-bar-fill" style={{width:Math.min(100,(p.v||0)/p.max*100)+'%',background:p.color}}/></div>}
              </div>
            ))}
          </div>
          {macros&&<div style={{fontSize:11,color:'var(--muted)',marginTop:6,lineHeight:1.5}}>
            Total batch: <strong>{macros.total}g</strong> · Fat: <strong>{macros.fat}g</strong> · Water: <strong>{macros.water}g</strong> · Sugar: <strong>{macros.sugar}g</strong> · Salt: <strong>{macros.saltG}g</strong>
            {data.macroCache&&<span style={{display:'block',marginTop:3}}>Flour equiv: <strong>{macros.flourEqG}g</strong> · Free water: <strong>{macros.freeWaterG}g</strong></span>}
          </div>}
          {(data.customParams||[]).map(cp=>(
            <div key={cp.id} style={{marginTop:8,padding:'6px 10px',border:'1px solid #b45309',borderRadius:6,background:'rgba(180,83,9,0.05)'}}>
              <div style={{display:'flex',justifyContent:'space-between',alignItems:'center'}}>
                <span style={{fontWeight:600,fontSize:12}}>{cp.label}: <strong>{cp.value}</strong>{cp.unit&&<span style={{marginLeft:3,fontSize:10}}>{cp.unit}</span>}</span>
                <button onClick={()=>removeCustomParam(cp.id)} style={{fontSize:10,background:'transparent',border:'none',cursor:'pointer',color:'var(--muted)'}}>&#x2715;</button>
              </div>
              {cp.explanation&&<div style={{fontSize:10,color:'var(--muted)',marginTop:2}}>{cp.explanation}</div>}
            </div>
          ))}
          <div style={{marginTop:8}}>
            {addingParam?(
              <div style={{display:'flex',gap:6,alignItems:'center'}}>
                <input value={newParamLabel} onChange={e=>setNewParamLabel(e.target.value)} placeholder="Parameter name..." style={{flex:1,fontSize:11,padding:'3px 6px',borderRadius:4,border:'1px solid var(--border)'}} onKeyDown={e=>{if(e.key==='Enter')addCustomParam();if(e.key==='Escape')setAddingParam(false)}}/>
                <button onClick={addCustomParam} disabled={analyzingCustom} style={{fontSize:10,padding:'3px 8px',borderRadius:4,border:'1px solid var(--id)',background:'transparent',color:'var(--id)',cursor:'pointer'}}>{analyzingCustom?'...':'\u2713'}</button>
                <button onClick={()=>setAddingParam(false)} style={{fontSize:10,padding:'3px 8px',borderRadius:4,border:'none',background:'transparent',color:'var(--muted)',cursor:'pointer'}}>&#x2715;</button>
              </div>
            ):(
              <button onClick={()=>setAddingParam(true)} style={{fontSize:10,padding:'2px 8px',borderRadius:4,border:'1px dashed var(--muted)',background:'transparent',color:'var(--muted)',cursor:'pointer'}}>&#x271A; Add parameter</button>
            )}
          </div>
        </div>}
      </div>
      {/* ─ Sensory ─ */}
      <div className="ID-section">
        <div className="ID-section-header" onClick={()=>toggleSection('sensory')}>
          <h3>👅 Sensory Evaluation</h3>
          <span style={{fontSize:11,color:'var(--id)'}}>{openSections.sensory?'▲':'▼'}</span>
        </div>
        {openSections.sensory&&<div className="ID-section-body">
          <div className="ID-sensory-grid">
            {SENSORY_ATTRS.map(attr=>{
              const val=data.sensory?.[attr]||{}
              return(
                <div className="ID-sensory-item" key={attr}>
                  <div className="ID-sensory-label">{attr}</div>
                  <div style={{fontSize:9.5,color:'var(--muted)',marginBottom:4}}>{SENSORY_LABELS[attr]}</div>
                  <div className="ID-star-row">
                    {[1,2,3,4,5].map(n=>(
                      <span key={n} className="ID-star" onClick={()=>setSensory(attr,'score',n)} style={{opacity:val.score>=n?1:.25}}>
                        {attr==='Overall score'?'⭐':'★'}
                      </span>
                    ))}
                    {val.score&&<span style={{fontFamily:'var(--mono)',fontSize:10,color:'var(--id)',marginLeft:4}}>{val.score}/5</span>}
                  </div>
                  <textarea className="ID-sensory-note" rows={2} value={val.note||''} onChange={e=>setSensory(attr,'note',e.target.value)} placeholder="Notes…"/>
                </div>
              )
            })}
          </div>
        </div>}
      </div>
      {/* ─ Timeline ─ */}
      <div className="ID-section">
        <div className="ID-section-header" onClick={()=>toggleSection('timeline')}>
          <h3>📅 Version Timeline</h3>
          <span style={{fontSize:11,color:'var(--id)'}}>{openSections.timeline?'▲':'▼'}</span>
        </div>
        {openSections.timeline&&<div className="ID-section-body">
          <div className="Q-field" style={{marginBottom:12}}>
            <label>Goal / target result</label>
            <textarea rows={2} value={data.goal||''} onChange={e=>setGoal(e.target.value)} placeholder="Describe the final desired result: color, texture, flavor profile, volume…" style={{width:'100%',border:'1px solid var(--rule)',borderRadius:6,padding:'8px 10px',fontSize:13,fontFamily:'var(--sans)',color:'var(--ink)',resize:'vertical',background:'#fff'}}/>
          </div>
          {(data.versions||[]).length===0&&<div style={{color:'var(--muted)',fontSize:12.5,marginBottom:12}}>No versions yet. Add one to start tracking your evolution.</div>}
          <div className="ID-timeline">
            {(data.versions||[]).map((v,i)=>(
              <div className="ID-timeline-item" key={v.id}>
                <div className={`ID-timeline-dot${i===(data.versions.length-1)?' active':''}${v.isGoal?' goal':''}`}/>
                <div className="ID-timeline-card">
                  {v.isGoal&&<span className="ID-timeline-badge goal">🎯 Target version</span>}
                  {!v.isGoal&&<span className="ID-timeline-badge version">v{i+1}</span>}
                  <div className="ID-timeline-date">{v.date}</div>
                  <div className="ID-timeline-title">{v.title}</div>
                  {v.notes&&<div className="ID-timeline-notes">{v.notes}</div>}
                  <button className="btn danger xs" style={{marginTop:6}} onClick={()=>removeVersion(v.id)}>Remove</button>
                </div>
              </div>
            ))}
          </div>
          {addingVersion?(
            <div style={{background:'#F8FBFF',border:'1px solid #C8DFF0',borderRadius:8,padding:12,marginTop:10}}>
              <div className="Q-field"><label>Version name / description</label><input value={newVersion.title} onChange={e=>setNewVersion(p=>({...p,title:e.target.value}))} placeholder="v3 — reduced sugar 5%, added orange zest"/></div>
              <div className="Q-field"><label>Notes</label><textarea rows={2} value={newVersion.notes} onChange={e=>setNewVersion(p=>({...p,notes:e.target.value}))} placeholder="What changed? Results?"/></div>
              <label style={{display:'flex',alignItems:'center',gap:6,fontSize:12.5,cursor:'pointer',marginBottom:10}}>
                <input type="checkbox" checked={newVersion.isGoal} onChange={e=>setNewVersion(p=>({...p,isGoal:e.target.checked}))}/> Mark as target version
              </label>
              <div style={{display:'flex',gap:7}}>
                <button className="btn id xs" onClick={addVersion}>Add</button>
                <button className="btn ghost xs" onClick={()=>setAddingVersion(false)}>Cancel</button>
              </div>
            </div>
          ):(
            <button className="btn id xs" style={{marginTop:10}} onClick={()=>setAddingVersion(true)}>＋ Add version</button>
          )}
        </div>}
      </div>
      {/* ─ Nutrition ─ */}
      <div className="ID-section">
        <div className="ID-section-header" onClick={()=>toggleSection('nutrition')}>
          <h3>🧬 Nutritional Table</h3>
          <span style={{fontSize:11,color:'var(--id)'}}>{openSections.nutrition?'▲':'▼'}</span>
          <span style={{fontFamily:'var(--mono)',fontSize:9,color:'var(--muted)',marginLeft:8}}>estimated</span>
        </div>
        {openSections.nutrition&&<div className="ID-section-body">
          <table className="ID-nutrition-table">
            <thead><tr><th>Nutrient</th><th>Per batch ({macros.total}g)</th><th>Per 100g</th></tr></thead>
            <tbody>
              {[
                {n:'Energy (kcal)',v:macros.cal,u:'kcal'},
                {n:'Fat',v:macros.fat,u:'g'},
                {n:'Water',v:macros.water,u:'g'},
                {n:'Carbohydrates',v:macros.carbs,u:'g'},
                {n:'of which sugars',v:macros.sugar,u:'g'},
                {n:'Protein',v:macros.protein,u:'g'},
                {n:'Salt',v:macros.saltG,u:'g'},
              ].map(r=>(
                <tr key={r.n}>
                  <td>{r.n}</td>
                  <td><strong>{r.v}</strong> {r.u}</td>
                  <td>{macros.total>0?Math.round(r.v/macros.total*1000)/10:0} {r.u}</td>
                </tr>
              ))}
            </tbody>
          </table>
          <div style={{fontSize:10,color:'var(--muted)',marginTop:8}}>* Estimated values based on ingredient type detection. Not for labeling use.</div>
        </div>}
      </div>
    </div>
  )
}

/* ── Recipe Comparison Panel ──────────────────────────────────────────────────────────── */

function ComparePanel({recipes,onClose}){
  const[sel,setSel]=useState([])
  const colors=['#1F3A4D','#BC6C2C','#2D6A4F','#5B3A8C','#1A6B6B']
  const selected=recipes.filter(r=>sel.includes(r.id))
  const allMacros=selected.map(r=>({r,m:calcMacros(r.ingredients)}))
  const PARAMS=[
    {k:'fatP',label:'Total fat %',max:50},
    {k:'hydration',label:'Hydration %',max:100},
    {k:'bakersHydration',label:"Baker's hydration %",max:120},
    {k:'sugarP',label:'Sugar %',max:40},
    {k:'saltP',label:"Salt (baker's %) %",max:4},
  ]
  return(
    <div className="Q-compare-overlay" onClick={e=>{if(e.target===e.currentTarget)onClose()}}>
      <div className="Q-compare-panel">
        <div className="Q-compare-header">
          <span style={{fontFamily:'var(--serif)',fontSize:17,color:'var(--navy)',flex:1}}>⚖ Recipe Comparison</span>
          <button className="btn ghost xs" onClick={onClose}>✕ Close</button>
        </div>
        <div className="Q-compare-body">
          <div style={{marginBottom:14}}>
            <div style={{fontFamily:'var(--mono)',fontSize:9.5,textTransform:'uppercase',letterSpacing:'.14em',color:'var(--muted)',marginBottom:7}}>Select recipes to compare</div>
            <div style={{display:'flex',flexWrap:'wrap',gap:6}}>
              {recipes.slice(0,20).map(r=>(
                <button key={r.id} onClick={()=>setSel(p=>p.includes(r.id)?p.filter(x=>x!==r.id):[...p,r.id])}
                  style={{fontSize:12,padding:'4px 10px',border:`1.5px solid ${sel.includes(r.id)?colors[sel.indexOf(r.id)%colors.length]:'var(--rule)'}`,borderRadius:20,background:sel.includes(r.id)?colors[sel.indexOf(r.id)%colors.length]+'22':'#fff',color:sel.includes(r.id)?colors[sel.indexOf(r.id)%colors.length]:'var(--muted)',cursor:'pointer',fontFamily:'var(--mono)',transition:'all .15s'}}>
                  {r.title}
                </button>
              ))}
            </div>
          </div>
          {selected.length>=2&&(
            <div>
              {/* Legend */}
              <div style={{display:'flex',flexWrap:'wrap',gap:10,marginBottom:14}}>
                {selected.map((r,i)=>(
                  <div key={r.id} style={{display:'flex',alignItems:'center',gap:6,fontSize:12}}>
                    <div style={{width:12,height:12,borderRadius:2,background:colors[i%colors.length]}}/>
                    {r.title}
                  </div>
                ))}
              </div>
              {PARAMS.map(param=>(
                <div key={param.k} style={{marginBottom:14}}>
                  <div style={{fontFamily:'var(--mono)',fontSize:9.5,textTransform:'uppercase',letterSpacing:'.12em',color:'var(--navy)',marginBottom:5}}>{param.label}</div>
                  {allMacros.map(({r,m},i)=>{
                    const v=m[param.k]||0;const pct=Math.min(100,v/param.max*100)
                    return(
                      <div key={r.id} style={{marginBottom:4}}>
                        <div style={{display:'flex',justifyContent:'space-between',fontSize:11,color:'var(--muted)',marginBottom:2}}>
                          <span>{r.title}</span>
                          <span style={{fontFamily:'var(--mono)',fontWeight:700,color:colors[i%colors.length]}}>{v}%</span>
                        </div>
                        <div className="ID-compare-bar">
                          <div className="ID-compare-fill" style={{width:pct+'%',background:colors[i%colors.length]}}/>
                        </div>
                      </div>
                    )
                  })}
                  {/* Diff annotation */}
                  {allMacros.length>=2&&(()=>{
                    const vals=allMacros.map(x=>x.m[param.k]||0)
                    const max=Math.max(...vals),min=Math.min(...vals),diff=Math.round((max-min)*10)/10
                    if(diff===0)return<div style={{fontSize:10,color:'var(--green)',marginTop:2}}>✓ Identical</div>
                    const maxR=allMacros[vals.indexOf(max)].r
                    return<div style={{fontSize:10,color:'var(--muted)',marginTop:2}}>↕ {diff}% difference · highest: <strong>{maxR.title}</strong></div>
                  })()}
                </div>
              ))}
              {/* Total grams */}
              <div style={{marginTop:16,borderTop:'1px solid var(--rule)',paddingTop:12}}>
                <div style={{fontFamily:'var(--mono)',fontSize:9.5,textTransform:'uppercase',letterSpacing:'.12em',color:'var(--navy)',marginBottom:8}}>Batch size</div>
                {allMacros.map(({r,m},i)=>(
                  <div key={r.id} style={{display:'flex',justifyContent:'space-between',padding:'4px 0',borderBottom:'1px dotted var(--rule)',fontSize:13}}>
                    <span style={{color:colors[i%colors.length],fontWeight:600}}>{r.title}</span>
                    <span style={{fontFamily:'var(--mono)'}}>{m.total}g</span>
                  </div>
                ))}
              </div>
            </div>
          )}
          {selected.length===1&&<div style={{color:'var(--muted)',fontSize:12.5,marginTop:16}}>Select at least one more recipe to compare.</div>}
          {selected.length===0&&<div style={{color:'var(--muted)',fontSize:12.5,marginTop:16}}>Select 2 or more recipes above to see a side-by-side comparison.</div>}
        </div>
      </div>
    </div>
  )
}

/* ── AI Assistant ──────────────────────────────────────────────────────────── */

const EXPIRY_LABELS=Object.keys(EXPIRY_OPTS)
function AIAssistant({recipe,onAction,onRequestSaveNote}){
  const[messages,setMessages]=useState([])
  const[input,setInput]=useState('')
  const[loading,setLoading]=useState(false)
  const[expiry,setExpiry]=useState('24 h')
  const endRef=useRef(null)
  const voice=useVoiceInput(t=>setInput(p=>p+(p?' ':'')+t),true)
  const CONV_KEY=`qdaid_conv_${recipe.id}`
  useEffect(()=>{try{const s=JSON.parse(localStorage.getItem(CONV_KEY)||'null');if(s&&s.exp>Date.now()){setMessages(s.msgs||[]);setExpiry(s.expiryLabel||'24 h')}}catch(_){}},[recipe.id])
  useEffect(()=>{if(!messages.length)return;try{localStorage.setItem(CONV_KEY,JSON.stringify({msgs:messages.map(m=>({role:m.role,content:m.content,clean:m.clean,hasActions:m.hasActions})),exp:Date.now()+(EXPIRY_OPTS[expiry]||EXPIRY_OPTS['24 h']),expiryLabel:expiry}))}catch(_){}},[messages,expiry,recipe.id])
  useEffect(()=>endRef.current?.scrollIntoView({behavior:'smooth'}),[messages])
  const chips=['Scale to 2 kg','Double recipe','Translate to Spanish','What is the hydration?','Best proofing temperature?','Add mixing tip']
  async function send(){
    if(!input.trim()||loading)return
    const msg={role:'user',content:input.trim()}
    const hist=[...messages,msg];setMessages(hist);setInput('');setLoading(true)
    try{
      const r=await askAssistant(hist,recipe)
      const text=r?.text||''
      const actionRx=/<ACTION>([\s\S]*?)<\/ACTION>/g;let m;const acts=[]
      while((m=actionRx.exec(text))!==null){try{acts.push(JSON.parse(m[1]))}catch{}}
      const clean=text.replace(/<ACTION>[\s\S]*?<\/ACTION>/g,'').trim()
      setMessages(p=>[...p,{role:'assistant',content:text,clean,hasActions:acts.length>0}])
      acts.forEach(a=>onAction(a))
    }catch(e){setMessages(p=>[...p,{role:'assistant',content:'Error: '+e.message,clean:'Error: '+e.message,hasActions:false}])}
    finally{setLoading(false)}
  }
  function clearConv(){setMessages([]);try{localStorage.removeItem(CONV_KEY)}catch(_){}}
  async function saveAsNote(){
    const content=[`AI Conversation — ${ts()}`,...messages.map(m=>`${m.role==='user'?'You':'AI'}: ${m.clean||m.content}`)].join('\n\n')
    onRequestSaveNote(content)
  }
  return(
    <div className="Q-assistant">
      <div className="Q-conv-toolbar">
        <span>Keep for:</span>
        <select value={expiry} onChange={e=>setExpiry(e.target.value)} style={{border:'1px solid var(--rule)',borderRadius:5,padding:'3px 6px',fontSize:11,fontFamily:'var(--mono)'}}>
          {EXPIRY_LABELS.map(l=><option key={l}>{l}</option>)}
        </select>
        {messages.length>0&&<><button className="btn xs ghost" onClick={saveAsNote}>Save as note</button><button className="btn xs ghost" onClick={clearConv}>Clear</button></>}
      </div>
      {messages.length===0&&(
        <div className="Q-assistant-welcome">
          <div style={{fontSize:32,marginBottom:8}}>🤖</div>
          <div style={{fontSize:14,fontWeight:700,color:'var(--ai)',marginBottom:5}}>AI Recipe Assistant</div>
          <div style={{fontSize:12.5,color:'var(--muted)',lineHeight:1.55,marginBottom:12}}>Ask anything or give instructions to modify this recipe.</div>
          <div className="Q-quick-chips">{chips.map(c=><button key={c} className="Q-chip" onClick={()=>setInput(c)}>{c}</button>)}</div>
        </div>
      )}
      <div className="Q-chat-wrap">
        {messages.map((m,i)=>(
          <div key={i} className={`Q-chat-msg ${m.role}`}>
            {m.clean||m.content}
            {m.hasActions&&<div className="Q-chat-action-badge">✓ Applied</div>}
          </div>
        ))}
        {loading&&<div className="Q-chat-msg assistant"><div className="Q-chat-typing"><span/><span/><span/></div></div>}
        <div ref={endRef}/>
      </div>
      <div className="Q-chat-input-area">
        <textarea className="Q-chat-input" value={input} onChange={e=>setInput(e.target.value)} rows={2} disabled={loading}
          placeholder="Ask or instruct… (Enter to send, Shift+Enter for newline)"
          onKeyDown={e=>{if(e.key==='Enter'&&!e.shiftKey){e.preventDefault();send()}}}/>
        <div style={{display:'flex',flexDirection:'column',gap:5}}>
          <button className={`Q-voice-btn${voice.recording?' recording':''}`} onClick={voice.recording?voice.stop:voice.start}>{voice.recording?'⏹':'🎙'}</button>
          <button className="btn ai xs" onClick={send} disabled={loading||!input.trim()}>↑</button>
        </div>
      </div>
    </div>
  )
}

/* ── App AI Chat ──────────────────────────────────────────────────────────── */

function parseAppActions(text){
  const rx=/<APP_ACTION>([\s\S]*?)<\/APP_ACTION>/g;let m;const acts=[]
  while((m=rx.exec(text))!==null){try{acts.push(JSON.parse(m[1]))}catch{}}
  return{clean:text.replace(/<APP_ACTION>[\s\S]*?<\/APP_ACTION>/g,'').trim(),actions:acts}
}
function AppAIChat({recipes,onAction,onClose}){
  const[messages,setMessages]=useState([])
  const[input,setInput]=useState('')
  const[loading,setLoading]=useState(false)
  const endRef=useRef(null)
  const voice=useVoiceInput(t=>setInput(p=>p+(p?' ':'')+t),true)
  useEffect(()=>endRef.current?.scrollIntoView({behavior:'smooth'}),[messages])
  const chips=['Add 3 French pastry recipes','Search for brioche','Delete the recipe named…','Create a recipe for sourdough bread','List all recipes with category']
  async function send(){
    if(!input.trim()||loading)return
    const msg={role:'user',content:input.trim()}
    const hist=[...messages,msg];setMessages(hist);setInput('');setLoading(true)
    try{
      const r=await askAppAssistant(hist,recipes)
      const{clean,actions}=parseAppActions(r?.text||'')
      setMessages(p=>[...p,{role:'assistant',content:r?.text||'',clean,hasActions:actions.length>0}])
      for(const act of actions)await onAction(act)
    }catch(e){setMessages(p=>[...p,{role:'assistant',content:'Error: '+e.message,clean:'Error: '+e.message,hasActions:false}])}
    finally{setLoading(false)}
  }
  return(
    <>
      <div className="Q-app-ai-header">
        <div className="Q-app-ai-title">🌐 App Assistant</div>
        <button className="btn ghost xs" onClick={onClose}>✕ Close</button>
      </div>
      <div className="Q-app-ai-msgs">
        {messages.length===0&&(
          <div style={{padding:'16px 0',textAlign:'center'}}>
            <div style={{fontSize:28,marginBottom:8}}>🌐</div>
            <div style={{fontSize:13,fontWeight:700,color:'var(--ai)',marginBottom:5}}>Quaderno AI+D — App Assistant</div>
            <div style={{fontSize:12,color:'var(--muted)',lineHeight:1.5,marginBottom:12}}>Search, create, batch-import, delete recipes. Use natural language.</div>
            <div className="Q-quick-chips" style={{justifyContent:'center'}}>{chips.map(c=><button key={c} className="Q-chip" onClick={()=>setInput(c)}>{c}</button>)}</div>
          </div>
        )}
        {messages.map((m,i)=>(
          <div key={i} className={`Q-chat-msg ${m.role}`} style={{maxWidth:'94%'}}>
            {m.clean||m.content}
            {m.hasActions&&<div className="Q-chat-action-badge">✓ Done</div>}
          </div>
        ))}
        {loading&&<div className="Q-chat-msg assistant"><div className="Q-chat-typing"><span/><span/><span/></div></div>}
        <div ref={endRef}/>
      </div>
      <div className="Q-app-ai-input">
        <textarea style={{flex:1,border:'1px solid var(--rule)',borderRadius:8,padding:'8px 10px',fontSize:13,fontFamily:'var(--sans)',color:'var(--ink)',resize:'none',background:'#fff'}}
          value={input} onChange={e=>setInput(e.target.value)} rows={2} disabled={loading}
          placeholder="Ask anything… (Enter to send)"
          onKeyDown={e=>{if(e.key==='Enter'&&!e.shiftKey){e.preventDefault();send()}}}/>
        <div style={{display:'flex',flexDirection:'column',gap:5}}>
          <button className={`Q-voice-btn${voice.recording?' recording':''}`} onClick={voice.recording?voice.stop:voice.start}>{voice.recording?'⏹':'🎙'}</button>
          <button className="btn ai xs" onClick={send} disabled={loading||!input.trim()}>↑</button>
        </div>
      </div>
    </>
  )
}

/* ── RecipeView ──────────────────────────────────────────────────────────── */

function RecipeView({recipe,onEdit,onDelete,onUpdate,allRecipes,onCopy}){
  const[tab,setTab]=useState('recipe')
  const[lightboxSrc,setLightboxSrc]=useState(null)
  const[checked,setChecked]=useState(new Set())
  const[highlightedSteps,setHighlightedSteps]=useState(new Set())
  const[showPct,setShowPct]=useState(false)
  const[pctMode,setPctMode]=useState('baker')
  const[pctBase,setPctBase]=useState('')
  const[showScale,setShowScale]=useState(false)
  const[scaleMode,setScaleMode]=useState('factor')
  const[scaleFactor,setScaleFactor]=useState('2')
  const[scalePieces,setScalePieces]=useState('')
  const[scaleGpp,setScaleGpp]=useState('')
  const[scaleTotal,setScaleTotal]=useState('')
  const[appliedScale,setAppliedScale]=useState(null)
  const[translating,setTranslating]=useState(false)
  const[translated,setTranslated]=useState(null)
  const[targetLang,setTargetLang]=useState('English')
  const[transErr,setTransErr]=useState('')
  const[exporting,setExporting]=useState(false)
  const[exportNotes,setExportNotes]=useState(false)
  const[showCopyLangMenu,setShowCopyLangMenu]=useState(false)
  const addNoteRef=useRef(null)
  useEffect(()=>{
    setChecked(new Set());setHighlightedSteps(new Set());setAppliedScale(null);setTranslated(null)
    setShowScale(false);setTab('recipe');setShowCopyLangMenu(false)
    // If this recipe has a fixed language, show banner
  },[recipe.id])
  const displayR=translated||recipe
  // Fix: always pass original thumbnail to exports
  const originalThumbnail=recipe.thumbnail
  const viewR=useMemo(()=>appliedScale?scaleRecipe(displayR,appliedScale.factor):displayR,[displayR,appliedScale])
  const sections=useMemo(()=>parseSections(viewR.ingredients||[]),[viewR])
  const totalGrams=useMemo(()=>getTotalGrams(viewR.ingredients||[]),[viewR])
  const pctOpts=useMemo(()=>({showPct,pctMode,pctBase,appliedScaleLabel:appliedScale?.label}),[showPct,pctMode,pctBase,appliedScale])
  function handleIngToggle(rawIdx){
    setChecked(prev=>{const next=new Set(prev);if(next.has(rawIdx))next.delete(rawIdx);else next.add(rawIdx);const names=[];(viewR.ingredients||[]).forEach((ing,i)=>{if(next.has(i)&&!/^##?\s+/.test(ing))names.push(parseIng(ing).name)});const steps=new Set();names.forEach(n=>findStepsForIng(n,viewR.steps||[]).forEach(i=>steps.add(i)));setHighlightedSteps(steps);return next})
  }
  function applyScale(){
    let factor=0,label=''
    if(scaleMode==='factor'){factor=parseFloat(scaleFactor)||0;if(!factor)return;label=`×${factor}`}
    else{const cur=getTotalGrams(recipe.ingredients||[]);if(!cur){alert('No gram quantities found. Use "× Multiply" mode.');return};let tg=0;if(scaleMode==='pieces'){const pc=parseFloat(scalePieces)||0,g=parseFloat(scaleGpp)||0;if(!pc||!g)return;tg=pc*g;label=`${pc}×${g}g=${tg.toFixed(0)}g`}else{tg=parseFloat(scaleTotal)||0;if(!tg)return;label=`${tg.toFixed(0)}g`};factor=tg/cur}
    setAppliedScale({factor,label});setShowScale(false);setChecked(new Set());setHighlightedSteps(new Set())
  }
  async function handleTranslate(){
    setTranslating(true);setTransErr('')
    try{
      const result=await translateRecipe(recipe,targetLang)
      // Preserve original thumbnail in translated version
      setTranslated({...result,thumbnail:recipe.thumbnail,source_photos:recipe.source_photos})
    }catch(e){setTransErr('Failed: '+e.message)}
    finally{setTranslating(false)}
  }
  async function handleCopyWithLang(lang){
    setShowCopyLangMenu(false)
    onCopy(recipe,lang||null)
  }
  async function handleAssistantAction(action){
    switch(action.type){
      case'scale':setAppliedScale({factor:action.factor,label:`AI ×${action.factor}`});break
      case'translate':setTranslating(true);try{const r=await translateRecipe(recipe,action.language);setTranslated({...r,thumbnail:recipe.thumbnail,source_photos:recipe.source_photos})}catch(e){setTransErr(e.message)}finally{setTranslating(false)};break
      case'update_field':await onUpdate({...recipe,[action.field]:action.value});break
      case'update_ingredients':await onUpdate({...recipe,ingredients:action.ingredients});break
      case'update_steps':await onUpdate({...recipe,steps:action.steps});break
      case'add_note':if(addNoteRef.current)addNoteRef.current(action.content);break
    }
  }
  async function handleRequestSaveNote(content){
    try{const tabs=parseTabs(recipe.notes_pad);const updated=tabs.map((t,i)=>i===0?{...t,content:t.content+(t.content?'\n\n':'')+content}:t);await onUpdate({...recipe,notes_pad:serializeTabs(updated)})}catch(e){console.error('Save note:',e)}
    setTab('notes')
  }
  async function handleSaveNotes(serialized){await onUpdate({...recipe,notes_pad:serialized})}
  async function handleSaveMedia(serialized){await onUpdate({...recipe,media_library:serialized})}
  async function handleSaveIdData(serialized){await onUpdate({...recipe,id_data:serialized})}
  const pctBaseOpts=useMemo(()=>(viewR.ingredients||[]).filter(i=>!/^##?\s+/.test(i)).map(i=>parseIng(i).name).filter((n,i,a)=>n&&a.indexOf(n)===i),[viewR])
  const recipeContent=(
    <div>
      <div className="Q-toolbar">
        {!appliedScale&&<button className={`btn xs ${showScale?'amber':'ghost'}`} onClick={()=>setShowScale(!showScale)}>⚖ Scale</button>}
        <button className={`btn xs ${showPct?'amber':'ghost'}`} onClick={()=>setShowPct(!showPct)}>% Baker's</button>
        <select style={{border:'1px solid var(--rule)',borderRadius:5,padding:'4px 7px',fontSize:11.5,fontFamily:'var(--mono)',background:'#fff',color:'var(--ink)'}} value={targetLang} onChange={e=>setTargetLang(e.target.value)}>
          {LANGS.map(l=><option key={l}>{l}</option>)}
        </select>
        <button className="btn xs green" onClick={handleTranslate} disabled={translating}>{translating?'…':`🌐 ${targetLang}`}</button>
        {transErr&&<span style={{color:'#9b2c2c',fontSize:10}}>{transErr}</span>}
        <div className="right">
          <button className="btn amber xs" onClick={()=>exportXLS(viewR,pctOpts)}>↓ XLS</button>
          <button className="btn amber xs" disabled={exporting} onClick={async()=>{setExporting(true);try{await exportImage(viewR,pctOpts,exportNotes,originalThumbnail)}finally{setExporting(false)}}}>↓ IMG</button>
          <button className="btn amber xs" onClick={()=>exportPDF(viewR,pctOpts,exportNotes,originalThumbnail)}>↓ PDF</button>
        </div>
      </div>
      <div className="Q-export-opts">
        <label><input type="checkbox" checked={exportNotes} onChange={e=>setExportNotes(e.target.checked)}/> Include notes in PDF/Image</label>
        <div style={{marginLeft:'auto',position:'relative'}}>
          <button className="btn teal xs" onClick={()=>setShowCopyLangMenu(p=>!p)}>📋 Copy recipe</button>
          {showCopyLangMenu&&(
            <div style={{position:'absolute',right:0,top:'100%',marginTop:4,background:'#fff',border:'1px solid var(--rule)',borderRadius:8,boxShadow:'0 4px 16px rgba(0,0,0,.12)',zIndex:50,minWidth:200,padding:8}}>
              <div style={{fontFamily:'var(--mono)',fontSize:9.5,textTransform:'uppercase',letterSpacing:'.12em',color:'var(--muted)',padding:'4px 8px 8px'}}>Copy as…</div>
              <button onClick={()=>handleCopyWithLang(null)} style={{width:'100%',textAlign:'left',padding:'7px 10px',background:'none',border:'none',cursor:'pointer',fontSize:13,borderBottom:'1px solid var(--rule)'}}>📋 Plain copy (same language)</button>
              {LANGS.map(l=>(
                <button key={l} onClick={()=>handleCopyWithLang(l)} style={{width:'100%',textAlign:'left',padding:'7px 10px',background:'none',border:'none',cursor:'pointer',fontSize:13,color:'var(--navy)'}}>🌐 {l} version (fixed)</button>
              ))}
            </div>
          )}
        </div>
      </div>
      {showScale&&!appliedScale&&(
        <div className="Q-scale-panel">
          <h4>Scale recipe</h4>
          <div style={{display:'flex',flexWrap:'wrap',gap:8,marginBottom:10}}>
            {[['factor','× Multiply'],['pieces','Pieces × g/piece'],['total','Total weight']].map(([k,l])=>(
              <label key={k} style={{display:'flex',alignItems:'center',gap:4,fontSize:12,cursor:'pointer'}}><input type="radio" checked={scaleMode===k} onChange={()=>setScaleMode(k)}/>{l}</label>
            ))}
          </div>
          {scaleMode==='factor'&&<div className="Q-scale-row"><label>Factor</label><input type="number" value={scaleFactor} onChange={e=>setScaleFactor(e.target.value)} placeholder="2" min=".01" step=".1"/><span style={{fontSize:11,color:'var(--muted)'}}>× all quantities</span></div>}
          {scaleMode==='pieces'&&<div className="Q-scale-row"><label>Pieces</label><input type="number" value={scalePieces} onChange={e=>setScalePieces(e.target.value)} placeholder="6"/><span style={{fontSize:11,color:'var(--muted)'}}>×</span><input type="number" value={scaleGpp} onChange={e=>setScaleGpp(e.target.value)} placeholder="1000"/><label>g/piece</label></div>}
          {scaleMode==='total'&&<div className="Q-scale-row"><label>Total</label><input type="number" value={scaleTotal} onChange={e=>setScaleTotal(e.target.value)} placeholder="2000"/><span style={{fontSize:11,color:'var(--muted)'}}>g · current: {getTotalGrams(recipe.ingredients||[]).toFixed(0)}g</span></div>}
          <div style={{display:'flex',gap:7}}><button className="btn amber xs" onClick={applyScale}>Apply</button><button className="btn ghost xs" onClick={()=>setShowScale(false)}>Cancel</button></div>
        </div>
      )}
      {showPct&&<div className="Q-pct-bar"><label>% Basis:</label><select value={pctMode} onChange={e=>setPctMode(e.target.value)}><option value="baker">Baker's % (flour=100%)</option><option value="mass">Total mass %</option><option value="custom">Custom base</option></select>{pctMode==='custom'&&<select value={pctBase} onChange={e=>setPctBase(e.target.value)}><option value="">— select —</option>{pctBaseOpts.map(n=><option key={n}>{n}</option>)}</select>}</div>}
      <div style={{fontFamily:'var(--mono)',fontSize:10,textTransform:'uppercase',letterSpacing:'.18em',color:'var(--navy)',marginBottom:7}}>
        Ingredients{checked.size>0&&<button style={{marginLeft:10,fontFamily:'var(--mono)',fontSize:9,background:'none',border:'none',cursor:'pointer',color:'var(--muted)',textDecoration:'underline'}} onClick={()=>{setChecked(new Set());setHighlightedSteps(new Set())}}>clear</button>}
      </div>
      {sections.map((sec,si)=>{
        const pctData=showPct?calcPct(sec.items,pctMode,pctBase):null
        const secG=sec.items.reduce((s,ing)=>{const p=parseIng(ing);return s+toGrams(p.qty,p.unit)},0)
        return(
          <div key={si}>
            {sec.name&&<div className="Q-sec-h"><span>{sec.name}</span></div>}
            <ul className="Q-ings">
              {sec.items.map((ing,ii)=>{
                const rawIdx=sec.rawIndices[ii],isCk=checked.has(rawIdx)
                const mm=String(ing).match(/^([\d.,]+\s*[^\s]+)\s{2,}(.+)$/)||String(ing).match(/^([\d.,]+\s*[a-zA-Z%]+)\s+(.+)$/)
                const pct=pctData?pctData[ii]:null
                return(
                  <li key={ii} className={`Q-ing-row${isCk?' checked':''}`} onClick={()=>handleIngToggle(rawIdx)}>
                    <span className="Q-ing-check">{isCk?'✓':'○'}</span>
                    {mm?<><span className="Q-ing-qty">{mm[1].trim()}</span><span className="Q-ing-name">{mm[2].trim()}</span></>:<span className="Q-ing-name" style={{flex:1}}>{ing}</span>}
                    {pct?.pct!=null&&<span className={`Q-pct-badge${pct.isBase?' base':''}`}>{pct.pct.toFixed(1)}%</span>}
                  </li>
                )
              })}
            </ul>
            {sec.name&&secG>0&&<div className="Q-subtotal">{sec.name}: {secG.toFixed(0)} g</div>}
          </div>
        )
      })}
      {totalGrams>0&&<div className="Q-grand-total">Total: {totalGrams.toFixed(0)} g</div>}
      {viewR.steps?.length>0&&<><div className="Q-steps-label">Method{highlightedSteps.size>0&&<span style={{fontFamily:'var(--mono)',fontSize:9,color:'var(--amber)',marginLeft:10}}>{highlightedSteps.size} step{highlightedSteps.size>1?'s':''} highlighted</span>}</div><ol className="Q-steps">{viewR.steps.map((s,i)=><li key={i} className={highlightedSteps.has(i)?'highlighted':''}>{s}</li>)}</ol></>}
      {viewR.notes&&<div className="Q-baker-note">{viewR.notes}</div>}
      {recipe.source_photos?.length>0&&<div style={{marginTop:16}}><div style={{fontFamily:'var(--mono)',fontSize:9.5,textTransform:'uppercase',letterSpacing:'.16em',color:'var(--muted)',marginBottom:7}}>Source photos</div><div style={{display:'flex',gap:7,flexWrap:'wrap'}}>{recipe.source_photos.map((src,i)=><img key={i} src={src} style={{height:64,borderRadius:5,cursor:'pointer',border:'1px solid var(--rule)'}} onClick={()=>setLightboxSrc(src)} alt=""/>)}</div></div>}
      <div className="Q-view-foot"><button className="btn" onClick={onEdit}>Edit</button><button className="btn danger" onClick={onDelete}>Delete</button></div>
    </div>
  )
  return(
    <div className="Q-view">
      <div className="Q-view-header">
        <h1>{viewR.title||'Untitled'}</h1>
        {recipe.thumbnail&&<img src={recipe.thumbnail} className="Q-recipe-thumb" onClick={()=>setLightboxSrc(recipe.thumbnail)} alt={recipe.title}/>}
      </div>
      {appliedScale&&<div className="Q-banner scale">⚖ Scaled — {appliedScale.label}<button onClick={()=>{setAppliedScale(null);setChecked(new Set());setHighlightedSteps(new Set())}}>Reset</button></div>}
      {translated&&<div className="Q-banner trans">🌐 {targetLang} translation (thumbnail preserved)<button onClick={()=>setTranslated(null)}>Original</button></div>}
      {recipe.fixed_lang&&<div className="Q-banner copy">📌 Fixed language version — {recipe.fixed_lang}</div>}
      {recipe.copied_from&&<div style={{fontFamily:'var(--mono)',fontSize:9.5,color:'var(--muted)',marginBottom:8}}>📋 Copy of: {allRecipes.find(r=>r.id===recipe.copied_from)?.title||recipe.copied_from}</div>}
      <dl className="Q-meta">
        {viewR.category&&<div className="Q-meta-item"><dt>Category</dt><dd>{viewR.category}</dd></div>}
        {viewR.time&&<div className="Q-meta-item"><dt>Time</dt><dd>{viewR.time}</dd></div>}
        {viewR.servings&&<div className="Q-meta-item"><dt>Yield</dt><dd>{viewR.servings}</dd></div>}
        {viewR.source&&<div className="Q-meta-item"><dt>Source</dt><dd>{viewR.source}</dd></div>}
      </dl>
      <div className="Q-tabs">
        {[['recipe','�� Recipe'],['notes','📝 Notes & Media'],['id','🔬 I+D'],['ai','🤖 AI']].map(([k,l])=>(
          <button key={k} className={`Q-tab-btn${tab===k?' active':''}${k==='ai'?' ai-tab':''}${k==='id'?' id-tab':''}`} onClick={()=>setTab(k)}>{l}</button>
        ))}
      </div>
      {tab==='recipe'&&recipeContent}
      {tab==='notes'&&<NotesPanel recipe={recipe} onSave={handleSaveNotes} onSaveMedia={handleSaveMedia} onAddNote={addNoteRef}/>}
      {tab==='id'&&<IDPanel recipe={recipe} onSave={handleSaveIdData}/>}
      {tab==='ai'&&<AIAssistant recipe={viewR} onAction={handleAssistantAction} onRequestSaveNote={handleRequestSaveNote}/>}
      {lightboxSrc&&<div className="Q-lightbox" onClick={()=>setLightboxSrc(null)}><img src={lightboxSrc} alt=""/></div>}
    </div>
  )
}

/* ── RecipeEditor ──────────────────────────────────────────────────────────── */

function RecipeEditor({initial,onSave,onCancel}){
  const initIngs=initial?.ingredients||[]
  const[r,setR]=useState(()=>({title:initial?.title||'',category:initial?.category||'',time:initial?.time||'',servings:initial?.servings||'',notes:initial?.notes||'',source:initial?.source||'Manual',notes_pad:initial?.notes_pad||'',thumbnail:initial?.thumbnail||'',source_photos:initial?.source_photos||[],steps:initial?.steps||[],id_data:initial?.id_data||'',media_library:initial?.media_library||'',fixed_lang:initial?.fixed_lang||null,copied_from:initial?.copied_from||null}))
  const[ingredientLines,setIngredientLines]=useState(initIngs)
  const[tab,setTab]=useState('text')
  const[images,setImages]=useState([])
  const[rawText,setRawText]=useState('')
  const[scanning,setScanning]=useState(false)
  const[err,setErr]=useState('')
  const[dragOver,setDragOver]=useState(false)
  const[lightboxSrc,setLightboxSrc]=useState(null)
  const set=k=>e=>setR(p=>({...p,[k]:e.target.value}))
  async function processFiles(files){setErr('');try{const imgs=Array.from(files).filter(f=>f.type.startsWith('image/'));if(!imgs.length){setErr('No image files.');return};const compressed=await Promise.all(imgs.map(f=>compressImage(f)));setImages(p=>[...p,...compressed])}catch(e){setErr('Image error: '+e.message)}}
  function handleFileInput(e){processFiles(e.target.files);e.target.value=''}
  function handleDrop(e){e.preventDefault();setDragOver(false);processFiles(e.dataTransfer.files)}
  useEffect(()=>{const onPaste=async e=>{if(tab!=='photo')return;const fs=Array.from(e.clipboardData?.items||[]).filter(i=>i.type.startsWith('image/')).map(i=>i.getAsFile()).filter(Boolean);if(fs.length){e.preventDefault();processFiles(fs)}};document.addEventListener('paste',onPaste);return()=>document.removeEventListener('paste',onPaste)},[tab])
  async function runFromPhotos(){if(!images.length){setErr('Add at least one photo.');return};setScanning(true);setErr('');try{const data=await extractWithClaude(images);setIngredientLines(data.ingredients||[]);setR(p=>({...p,title:data.title||p.title,category:data.category||p.category,time:data.time||p.time,servings:data.servings||p.servings,notes:data.notes||p.notes,source:'Photo',source_photos:images.map(im=>im.url),steps:data.steps||p.steps}))}catch(e){setErr('Could not read photos. ('+e.message+')')}finally{setScanning(false)}}
  async function runFromText(){if(!rawText.trim()){setErr('Paste or type the recipe first.');return};setScanning(true);setErr('');try{const data=await structureText(rawText);setIngredientLines(data.ingredients||[]);setR(p=>({...p,title:data.title||p.title,category:data.category||p.category,time:data.time||p.time,servings:data.servings||p.servings,notes:data.notes||p.notes,source:'Text',steps:data.steps||p.steps}));setRawText('')}catch(e){setErr('Could not structure text. ('+e.message+')')}finally{setScanning(false)}}
  async function handleThumbnail(e){const f=e.target.files?.[0];if(!f)return;try{const d=await compressThumbnail(f);setR(p=>({...p,thumbnail:d}))}catch(e){setErr('Thumbnail error: '+e.message)};e.target.value=''}
  function save(){onSave({id:initial?.id,title:r.title.trim()||'Untitled',category:r.category.trim(),time:r.time.trim(),servings:r.servings.trim(),notes:r.notes.trim(),source:r.source||'Manual',notes_pad:r.notes_pad||'',thumbnail:r.thumbnail||'',source_photos:r.source_photos||[],ingredients:ingredientLines.filter(Boolean),steps:r.steps||[],id_data:r.id_data||'',media_library:r.media_library||'',fixed_lang:r.fixed_lang||null,copied_from:r.copied_from||null,createdAt:initial?.createdAt||Date.now()})}
  return(
    <div className="Q-ed">
      <h2>{initial?.id?'Edit recipe':'New recipe'}</h2>
      <div style={{border:'1.5px solid var(--rule)',borderRadius:10,marginBottom:18,overflow:'hidden'}}>
        <div style={{display:'flex',borderBottom:'1px solid var(--rule)',background:'#f5efe6'}}>
          {[['text','📋 Paste text'],['photo','📷 From photo']].map(([k,l])=>(
            <button key={k} onClick={()=>{setTab(k);setErr('')}} style={{flex:1,padding:'9px 8px',border:'none',cursor:'pointer',fontFamily:'var(--mono)',fontSize:10.5,fontWeight:600,textTransform:'uppercase',letterSpacing:'.1em',background:tab===k?'#fff':'transparent',color:tab===k?'var(--navy)':'var(--muted)',borderBottom:tab===k?'2px solid var(--amber)':'2px solid transparent'}}>{l}</button>
          ))}
        </div>
        <div style={{padding:'13px 15px'}}>
          {tab==='text'&&<><p style={{fontSize:11.5,color:'var(--muted)',margin:'0 0 9px',lineHeight:1.5}}>Paste any recipe text. Claude structures it automatically.</p><textarea value={rawText} onChange={e=>setRawText(e.target.value)} rows={6} placeholder="Paste recipe text here…" style={{width:'100%',border:'1px solid var(--rule)',borderRadius:7,padding:'9px 11px',fontSize:13,fontFamily:'var(--sans)',color:'var(--ink)',resize:'vertical',background:'#fff',display:'block',marginBottom:9}}/><button className="btn amber xs" onClick={runFromText} disabled={scanning||!rawText.trim()} style={{width:'100%',padding:'9px',fontSize:13}}>{scanning?'Structuring…':'Structure with Claude →'}</button></>}
          {tab==='photo'&&<><p style={{fontSize:11.5,color:'var(--muted)',margin:'0 0 9px',lineHeight:1.5}}>Upload 1–6 photos. Auto-compressed.</p><div onDrop={handleDrop} onDragOver={e=>{e.preventDefault();setDragOver(true)}} onDragLeave={()=>setDragOver(false)} style={{position:'relative',borderRadius:8,marginBottom:9,border:`2px dashed ${dragOver?'var(--navy)':'var(--amber)'}`,background:dragOver?'#EAF2EE':'rgba(188,108,44,.05)',padding:'18px 12px',textAlign:'center',cursor:scanning?'default':'pointer'}}><div style={{pointerEvents:'none'}}><div style={{fontSize:22,marginBottom:4}}>📷</div><div style={{fontSize:12,fontWeight:600,color:'var(--navy)'}}>Tap · drag & drop · paste ⌘V</div></div><input type="file" accept="image/*" multiple disabled={scanning} onChange={handleFileInput} style={{position:'absolute',inset:0,opacity:0,width:'100%',height:'100%',cursor:scanning?'default':'pointer'}}/></div>{images.length>0&&<><div className="Q-thumbs">{images.map((im,i)=><div className="Q-thumb" key={i}><img src={im.url} alt=""/><button onClick={()=>setImages(p=>p.filter((_,j)=>j!==i))} disabled={scanning}>×</button></div>)}</div><button className="btn amber xs" onClick={runFromPhotos} disabled={scanning} style={{marginTop:7,width:'100%',padding:'9px',fontSize:13}}>{scanning?`Reading ${images.length} photo${images.length>1?'s':''}…`:'Extract with Claude →'}</button></>}}</>
          }
          {err&&<div className="Q-err" style={{marginTop:7}}>{err}</div>}
        </div>
      </div>
      <div className="Q-field">
        <label>Thumbnail photo</label>
        <div style={{display:'flex',gap:10,alignItems:'center',flexWrap:'wrap'}}>
          {r.thumbnail&&<img src={r.thumbnail} style={{width:58,height:58,borderRadius:6,objectFit:'cover',cursor:'pointer',border:'1px solid var(--rule)'}} onClick={()=>setLightboxSrc(r.thumbnail)} alt=""/>}
          <label className="btn ghost xs" style={{cursor:'pointer'}}>{r.thumbnail?'Change':'Add photo'}<input type="file" accept="image/*" onChange={handleThumbnail} style={{display:'none'}}/></label>
          {r.thumbnail&&<button className="btn danger xs" onClick={()=>setR(p=>({...p,thumbnail:''}))}>Remove</button>}
        </div>
        <div className="hint">Compressed · appears in list, recipe header, and all exports</div>
      </div>
      <div className="Q-field"><label>Title</label><input value={r.title} onChange={set('title')} placeholder="Panettone Classico"/></div>
      <div className="Q-grid2">
        <div className="Q-field"><label>Category</label><input value={r.category} onChange={set('category')} placeholder="Grandi Lievitati"/></div>
        <div className="Q-field"><label>Source</label><input value={r.source} onChange={set('source')}/></div>
      </div>
      <div className="Q-grid2">
        <div className="Q-field"><label>Time</label><input value={r.time} onChange={set('time')} placeholder="~36 h"/></div>
        <div className="Q-field"><label>Yield</label><input value={r.servings} onChange={set('servings')} placeholder="2 × 1 kg"/></div>
      </div>
      <div className="Q-field">
        <label>Ingredients — drag ⠇ to reorder</label>
        <DraggableIngList lines={ingredientLines} onChange={setIngredientLines}/>
        <div className="hint">Drag to reorder · × to remove · use <strong>+ Section</strong> for ## headers</div>
      </div>
      <div className="Q-field">
        <label>Method — one step per line</label>
        <textarea rows={7} value={(r.steps||[]).join('\n')} onChange={e=>setR(p=>({...p,steps:e.target.value.split('\n')}))} placeholder="Step 1…"/>
      </div>
      <div className="Q-field"><label>Baker's notes</label><textarea rows={2} value={r.notes} onChange={set('notes')} placeholder="Temperatures, flour specs, adjustments…"/></div>
      {r.fixed_lang&&<div style={{fontFamily:'var(--mono)',fontSize:10,color:'var(--teal)',marginBottom:10}}>📌 Fixed language: {r.fixed_lang}</div>}
      <div className="Q-ed-foot">
        <button className="btn" onClick={save}>Save recipe</button>
        <button className="btn ghost" onClick={onCancel}>Cancel</button>
      </div>
      {lightboxSrc&&<div className="Q-lightbox" onClick={()=>setLightboxSrc(null)}><img src={lightboxSrc} alt=""/></div>}
    </div>
  )
}

/* ── App ──────────────────────────────────────────────────────────── */

export default function App(){
  const[recipes,setRecipes]=useState([])
  const[loading,setLoading]=useState(true)
  const[selId,setSelId]=useState(null)
  const[mode,setMode]=useState('view')
  const[q,setQ]=useState('')
  const[saveErr,setSaveErr]=useState('')
  const[showAppAI,setShowAppAI]=useState(false)
  const[showCompare,setShowCompare]=useState(false)
  useEffect(()=>{dbLoad().then(data=>{setRecipes(data);if(data[0])setSelId(data[0].id)}).catch(e=>setSaveErr('Load failed: '+e.message)).finally(()=>setLoading(false))},[])
  async function saveRecipe(rec){
    try{
      const saved=rec.id&&recipes.some(x=>x.id===rec.id)?await dbUpdate(rec):await dbInsert(rec)
      setRecipes(p=>{const ex=p.some(x=>x.id===saved.id);return ex?p.map(x=>x.id===saved.id?saved:x):[saved,...p]})
      setSelId(saved.id);setMode('view')
    }catch(e){setSaveErr('Save failed: '+e.message)}
  }
  async function updateRecipe(updated){
    try{const saved=await dbUpdate(updated);setRecipes(p=>p.map(x=>x.id===saved.id?saved:x))}
    catch(e){setSaveErr('Update failed: '+e.message)}
  }
  async function deleteRecipe(id){
    if(!window.confirm('Delete this recipe?'))return
    try{await dbDelete(id);const next=recipes.filter(x=>x.id!==id);setRecipes(next);setSelId(next[0]?.id||null);setMode('view')}
    catch(e){setSaveErr('Delete failed: '+e.message)}
  }
  async function copyRecipe(sourceRecipe,fixedLang){
    try{
      let rec={...sourceRecipe,id:undefined,title:sourceRecipe.title+(fixedLang?` (${fixedLang})`:'  (Copy)'),notes_pad:'',media_library:'',id_data:'',fixed_lang:fixedLang||null,copied_from:sourceRecipe.id}
      // If fixedLang requested, translate first
      if(fixedLang){
        try{
          const translated=await translateRecipe(sourceRecipe,fixedLang)
          rec={...rec,...translated,thumbnail:sourceRecipe.thumbnail,source_photos:sourceRecipe.source_photos,fixed_lang:fixedLang,copied_from:sourceRecipe.id}
        }catch(e){console.warn('Translation failed, copying as-is',e)}
      }
      const saved=await dbInsert(rec)
      setRecipes(p=>[saved,...p])
      setSelId(saved.id);setMode('view')
    }catch(e){setSaveErr('Copy failed: '+e.message)}
  }
  async function handleAppAIAction(action){
    switch(action.type){
      case'create_recipe':try{const saved=await dbInsert({...action.recipe,notes_pad:'',thumbnail:'',source_photos:[],id_data:'',media_library:'',fixed_lang:null,copied_from:null});setRecipes(p=>[saved,...p])}catch(e){setSaveErr('Create failed: '+e.message)};break
      case'batch_create':try{const created=await Promise.all((action.recipes||[]).map(r=>dbInsert({...r,notes_pad:'',thumbnail:'',source_photos:[],id_data:'',media_library:'',fixed_lang:null,copied_from:null})));setRecipes(p=>[...created,...p])}catch(e){setSaveErr('Batch create failed: '+e.message)};break
      case'delete_recipe':if(window.confirm(`Delete "${action.title||action.id}"?`)){try{await dbDelete(action.id);setRecipes(p=>p.filter(r=>r.id!==action.id));if(selId===action.id)setSelId(null)}catch(e){setSaveErr('Delete failed: '+e.message)}};break
      case'select_recipe':setSelId(action.id);setMode('view');setShowAppAI(false);break
      case'search':setQ(action.query||'');setShowAppAI(false);break
    }
  }
  const sel=recipes.find(x=>x.id===selId)||null
  const filtered=recipes.filter(r=>{
    if(!q.trim())return true
    return[r.title,r.category,...(r.ingredients||[])].join(' ').toLowerCase().includes(q.toLowerCase())
  })
  const isOpen=mode!=='view'||!!sel
  return(
    <div className="Q" data-open={isOpen?'1':'0'}>
      <style>{CSS}</style>
      <header className="Q-top">
        <div className="Q-brand">
          Quaderno
          <span className="ai-badge">AI</span>
          <span className="id-badge">+D</span>
        </div>
        <div className="Q-top-right">
          {saveErr&&<span style={{color:'#9b2c2c',fontSize:10,maxWidth:160,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{saveErr}</span>}
          <span style={{fontFamily:'var(--mono)',fontSize:10,color:'var(--muted)'}}>{!loading&&`${recipes.length} recipe${recipes.length!==1?'s':''}`}</span>
          <button className="btn id xs" onClick={()=>setShowCompare(true)} title="Compare recipes">⚖ Compare</button>
          <button className="btn ai xs" onClick={()=>setShowAppAI(true)} title="App AI Assistant">🌐 AI</button>
          <button className="btn amber" onClick={()=>{setMode('new');setSelId(null)}}>＋ New</button>
        </div>
      </header>
      <div className="Q-body">
        <aside className="Q-side">
          <div className="Q-search"><input value={q} onChange={e=>setQ(e.target.value)} placeholder="Search recipes…"/></div>
          <div className="Q-list">
            {loading&&<div className="Q-msg">Loading…</div>}
            {!loading&&!filtered.length&&<div className="Q-msg">{q?'No matches.':'No recipes yet!'}</div>}
            {filtered.map(r=>(
              <button key={r.id} className="Q-list-item" aria-selected={r.id===selId&&mode==='view'} onClick={()=>{setSelId(r.id);setMode('view')}}>
                {r.thumbnail?<img src={r.thumbnail} className="Q-list-thumb" alt=""/>:<div className="Q-list-thumb-ph">🍞</div>}
                <div>
                  <h4>{r.title}</h4>
                  <span>
                    {[r.category,r.source].filter(Boolean).join(' · ')||'—'}
                    {r.fixed_lang&&` · 📌${r.fixed_lang}`}
                  </span>
                </div>
              </button>
            ))}
          </div>
        </aside>
        <main className="Q-main">
          <div className="Q-pane">
            <button className="btn ghost xs Q-back-btn" style={{marginBottom:14}} onClick={()=>{setMode('view');setSelId(null)}}>← All recipes</button>
            {mode==='new'&&<RecipeEditor onSave={saveRecipe} onCancel={()=>{setMode('view');setSelId(recipes[0]?.id||null)}}/> }
            {mode==='edit'&&sel&&<RecipeEditor initial={sel} onSave={saveRecipe} onCancel={()=>setMode('view')}/>}
            {mode==='view'&&sel&&<RecipeView key={sel.id} recipe={sel} onEdit={()=>setMode('edit')} onDelete={()=>deleteRecipe(sel.id)} onUpdate={updateRecipe} allRecipes={recipes} onCopy={copyRecipe}/>}
            {mode==='view'&&!sel&&!loading&&(
              <div className="Q-hero">
                <div className="glyph">❦</div>
                <h2>Quaderno AI+D</h2>
                <p>Professional recipe intelligence with R&D tools. Baker's percentages, sensory evaluation, version tracking, media library, and AI assistance — all in one place.</p>
                <button className="btn amber" onClick={()=>setMode('new')}>Add first recipe</button>
              </div>
            )}
          </div>
        </main>
      </div>
      {showAppAI&&(
        <div className="Q-app-ai-overlay" onClick={e=>{if(e.target===e.currentTarget)setShowAppAI(false)}}>
          <div className="Q-app-ai-panel">
            <AppAIChat recipes={recipes} onAction={handleAppAIAction} onClose={()=>setShowAppAI(false)}/>
          </div>
        </div>
      )}
      {showCompare&&<ComparePanel recipes={recipes} onClose={()=>setShowCompare(false)}/>}
    </div>
  )
}
