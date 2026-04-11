/**
 * STM Bot Engine V2.0 - Súper Buscador Centralizado
 * Se recomienda cargar este script justo antes del cierre de </body>
 */

(function () {
    // 1. Base de Conocimientos Estática (Módulos de la App)
    const sitePages = [
        { t: "Actividades STM", d: "Beneficios de temporada y eventos", f: "actividades.html", tags: "libros bolsa tiles guardapolvos escolares mochilas canasta educacion hijos accion social escolaridad beneficio primaria secundaria", ico: "fa-calendar-check" },
        { t: "Reclamo Ganancias", d: "Información judicial y trámites", f: "reclamoganancias.html", tags: "ganancias impuesto afip arca ips documentos cronograma", ico: "fa-scale-balanced" },
        { t: "Reclamo Antigüedad", d: "Reclamos judiciales 1995-2014", f: "reclamoantiguedad.html", tags: "antiguedad reclamo judicial abogados demanda activos jubilados", ico: "fa-clock-rotate-left" },
        { t: "Servicios STM", d: "Beneficios, camping, subsidios y más", f: "servicios.html", tags: "camping quinchos piletas préstamos proveeduría casamiento nacimiento útiles libros guardapolvos farmacia legal asesoría", ico: "fa-screwdriver-wrench" },
        { t: "Convenios Especiales", d: "Descuentos en comercios y salud", f: "convenios.html", tags: "descuentos convenios beneficios automotor salud veterinaria recreacion comercios", ico: "fa-handshake" },
        { t: "Autoridades", d: "Cuerpo directivo y secretarías", f: "autoridades.html", tags: "quien mando jefe directiva secretario conducción", ico: "fa-users-gear" },
        { t: "Afiliación Online", d: "Inicia tu trámite para sumarte al sindicato", f: "afiliacion.html", tags: "alta inscripción socio sindicato ficha nueva sumarse", ico: "fa-user-plus" },
        { t: "Consulta de Trámite", d: "Estado de tus gestiones y expedientes", f: "consulta.html", tags: "ver estado proceso seguimiento expediente trámite", ico: "fa-magnifying-glass-chart" },
        { t: "Contactos e Internos", d: "Directorio de internos y secretarías", f: "contactos.html", tags: "teléfono oficina directivos secretario ayuda interno llamar", ico: "fa-address-book" },
        { t: "Género y Diversidad", d: "Protección y bienestar integral", f: "generoydiversidad.html", tags: "ayuda derechos protección género diversidad violencia bienestar", ico: "fa-hand-holding-heart" },
        { t: "Mini Turismo", d: "Viajes y escapadas para afiliados", f: "miniturismo.html", tags: "viaje turismo miramar escapada fin de semana costa hotel", ico: "fa-umbrella-beach" },
        { t: "Capacitación", d: "Formación y becas para estudiantes", f: "capacitacion.html", tags: "estudio terciario universitario beca kit utiles voucher formacion", ico: "fa-graduation-cap" },
        { t: "Novedades STM", d: "Últimas noticias y comunicados", f: "novedades.html", tags: "noticias novedad aviso boletin comunicado sorteos informacion ultimo", ico: "fa-newspaper" },
        { t: "Inicio", d: "Panel principal de opciones", f: "opciones.html", tags: "inicio dashboard panel principal volver", ico: "fa-house" }
    ];

    // 2. Base de Conocimientos Dinámica
    const BOT_SHEET_ID = "16Y723omF3l38Ntq0MUSh-ZMYnRIvYmctrscidox5ktc";
    let botNews = [];
    let botConvenios = [];

    async function fetchExternalData() {
        // Fetch Novedades GID: 195669740
        if(botNews.length === 0) {
            try {
                const res = await fetch(`https://docs.google.com/spreadsheets/d/${BOT_SHEET_ID}/gviz/tq?tqx=out:json&gid=195669740`);
                const text = await res.text();
                const json = JSON.parse(text.substring(text.indexOf("{"), text.lastIndexOf("}") + 1));
                botNews = (json.table.rows || []).map(r => ({
                    titulo: (r.c[1]?.v || "").toString(),
                    desc:   (r.c[2]?.v || "").toString()
                }));
            } catch (e) { console.warn("Bot: Error cargando Novedades dinámicas", e); }
        }
        
        // Fetch Convenios GID: 1025588963
        if(botConvenios.length === 0) {
            try {
                const res = await fetch(`https://docs.google.com/spreadsheets/d/${BOT_SHEET_ID}/gviz/tq?tqx=out:json&gid=1025588963`);
                const text = await res.text();
                const json = JSON.parse(text.substring(text.indexOf("{"), text.lastIndexOf("}") + 1));
                botConvenios = (json.table.rows || []).map(r => ({
                    rubro: (r.c[1]?.v || "").toString(),
                    titulo: (r.c[2]?.v || "").toString(),
                    desc: (r.c[3]?.v || "").toString(),
                    activo: (r.c[4]?.v || "").toString().trim().toUpperCase() === "SI"
                })).filter(c => c.activo);
            } catch (e) { console.warn("Bot: Error cargando Convenios dinámicos", e); }
        }
    }

    function normalizeStr(str) { return (str || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[¿?¡!.,;:()"'\-\/]/g, "").toLowerCase(); }

    // Algoritmo de distancia para tolerar errores ortográficos (tipeos)
    function levenshtein(a, b) {
        if(a.length === 0) return b.length;
        if(b.length === 0) return a.length;
        const matrix = [];
        for (let i = 0; i <= b.length; i++) matrix[i] = [i];
        for (let j = 0; j <= a.length; j++) matrix[0][j] = j;
        for (let i = 1; i <= b.length; i++) {
            for (let j = 1; j <= a.length; j++) {
                if (b.charAt(i - 1) === a.charAt(j - 1)) {
                    matrix[i][j] = matrix[i - 1][j - 1];
                } else {
                    matrix[i][j] = Math.min(matrix[i - 1][j - 1] + 1, Math.min(matrix[i][j - 1] + 1, matrix[i - 1][j] + 1));
                }
            }
        }
        return matrix[b.length][a.length];
    }

    function isFuzzyMatch(word, target) {
        if (target.includes(word)) return true;
        
        // NLP natural root extraction (Stemming para español)
        let root = word.replace(/(miento|amiento|imiento|cion|idad|eria|mente|ando|iendo|as|os|es|s)$/, '');
        if (root.length >= 4 && target.includes(root)) return true;

        if (word.length < 4) return false;
        
        const targetWords = target.split(/\s+/);
        for (let tw of targetWords) {
            if(tw.length < 4) continue;
            // Inclusión inversa (ej: si el usuario escribe asesoramiento y el target solo tiene asesor)
            if(word.includes(tw) && tw.length >= 5) return true;
            // 1 typo para palabras cortas (4-5), 2 typos para 6+ letras
            const maxTypos = word.length >= 6 ? 2 : 1;
            if (levenshtein(word, tw) <= maxTypos) return true;
        }
        return false;
    }

    // Interceptar la función global
    window.toggleBot = function (show) {
        const modal = document.getElementById('modalBot');
        if (modal) {
            modal.classList.toggle('active', show);
            if (show) {
                const input = document.getElementById('botSearch');
                if (input) input.focus();
                fetchExternalData(); // Pre-cardar datos al abrir el bot
            }
        }
    };
    
    // Si la página se llama opciones.html interceptamos también toggleModal por las dudas
    const originalToggleModal = window.toggleModal;
    window.toggleModal = function(show) {
        if(document.getElementById('modalBot')) {
            window.toggleBot(show);
        } else if(originalToggleModal) {
            originalToggleModal(show);
        }
    }

    window.addEventListener('DOMContentLoaded', () => {
        // Auto-detectar página actual y marcar botón activo en la botonera
        const currentPage = window.location.pathname.split('/').pop() || 'index.html';
        const navBtns = document.querySelectorAll('.bottom-nav .nav-btn');
        navBtns.forEach(btn => {
            const href = (btn.getAttribute('href') || '').split('/').pop();
            if (href && href !== '#' && currentPage === href) {
                btn.classList.add('active');
                btn.classList.remove('nav-home-pulse'); // No parpadear si ya está activo
            }
        });

        let oldInput = document.getElementById('botSearch');
        
        if (oldInput) {
            // Remueve listeners antiguos sobreescribiendo el clon
            let newInput = oldInput.cloneNode(true);
            oldInput.parentNode.replaceChild(newInput, oldInput);

            let searchTimer = null;
            newInput.addEventListener('input', (e) => {
                const val = normalizeStr(e.target.value).trim();
                const res = document.getElementById('botRes');
                if (searchTimer) clearTimeout(searchTimer);

                if (val.length <= 2) {
                    res.innerHTML = "";
                    return;
                }

                res.innerHTML = `<div style="padding:20px; text-align:center; color:#94a3b8; font-size:0.85rem;" class="animate__animated animate__pulse animate__infinite">Buscando con el motor STM...</div>`;

                searchTimer = setTimeout(() => {
                    res.innerHTML = "";
                    const rawWords = val.split(" ").filter(w => w.length > 2);
                    const stopWords = ["quiero","necesito","busco","ver","saber","como","donde","de","la","el","los","las","un","una","unos","unas","mi","tu","su","para","con","por","en","hola","bot","stm","que","cual","quien","porfavor","favor","cuando","sale","hay","tiene","tienen","puedo","puede","hacer","del","al","sobre","mas","esta","estan","son","ser","era","fue","muy","tambien","algo","eso","ese","esa","esto","este","esta","nos","les","les","sin","pero","porque","desde","hasta","entre","otro","otra","otros","cada","todo","toda","todos","todas"];
                    const words = rawWords.filter(w => !stopWords.includes(w));
                    
                    // Si el usuario escribió puros conectores ("hola busco"), no filtramos todo
                    const searchTerms = words.length > 0 ? words : rawWords;
                    
                    // Match sitepages
                    const matchedPages = sitePages.filter(p => {
                        const target = normalizeStr(p.t + " " + p.tags + " " + p.d);
                        return searchTerms.every(w => isFuzzyMatch(w, target)) || isFuzzyMatch(val, target);
                    });
                    
                    // Match Convenios
                    const matchedConv = botConvenios.filter(c => {
                        const target = normalizeStr(c.rubro + " " + c.titulo + " " + c.desc);
                        return searchTerms.every(w => isFuzzyMatch(w, target)) || isFuzzyMatch(val, target);
                    });

                    // Match Novedades
                    const matchedNews = botNews.filter(n => {
                        const target = normalizeStr(n.titulo + " " + n.desc);
                        return searchTerms.every(w => isFuzzyMatch(w, target)) || isFuzzyMatch(val, target);
                    });

                    if (matchedPages.length === 0 && matchedConv.length === 0 && matchedNews.length === 0) {
                        res.innerHTML = `<div style="padding:20px; text-align:center; color:#94a3b8; font-size:0.85rem;">No encontré resultados. Intenta con otra palabra clave.</div>`;
                        return;
                    }

                    // Render Pages
                    if (matchedPages.length > 0) {
                        res.innerHTML += `<div class="bot-results-title" style="font-size: 0.72rem; font-weight: 800; color: var(--stm-primary); text-transform: uppercase; letter-spacing: 0.5px; margin: 15px 0 10px 5px;">MÓDULOS DE LA PLATAFORMA</div>`;
                        matchedPages.forEach(p => {
                            res.innerHTML += `<a href="${p.f}" class="res-item-bot" style="display:flex; align-items:center; gap:12px; padding:12px; border-radius:15px; margin-bottom:8px; text-decoration:none; color:inherit; background:#f8fafc; border:1px solid #f1f5f9; display:block">
                                <div style="display:flex; align-items:center; gap:10px">
                                    <div style="width:30px; height:30px; border-radius:8px; background:var(--stm-primary); color:white; display:flex; align-items:center; justify-content:center; font-size:0.8rem"><i class="fa-solid ${p.ico}"></i></div>
                                    <div><h4 style="margin:0; font-size:0.85rem">${p.t}</h4><p style="margin:0; font-size:0.7rem; opacity:0.7">${p.d}</p></div>
                                </div>
                            </a>`;
                        });
                    }

                    // Render Convenios
                    if (matchedConv.length > 0) {
                        res.innerHTML += `<div class="bot-results-title" style="font-size: 0.72rem; font-weight: 800; color: #10b981; text-transform: uppercase; letter-spacing: 0.5px; margin: 15px 0 10px 5px;">CONVENIOS ENCONTRADOS</div>`;
                        matchedConv.slice(0, 5).forEach(c => {
                            res.innerHTML += `<a href="convenios.html" class="res-item-bot" style="display:flex; align-items:center; gap:12px; padding:12px; border-radius:15px; margin-bottom:8px; text-decoration:none; color:inherit; background:#ecfdf5; border:1px solid #d1fae5; display:block">
                                <div style="display:flex; align-items:center; gap:10px">
                                    <div style="width:30px; height:30px; border-radius:8px; background:#10b981; color:white; display:flex; align-items:center; justify-content:center; font-size:0.8rem"><i class="fa-solid fa-handshake"></i></div>
                                    <div><h4 style="margin:0; font-size:0.85rem">${c.titulo}</h4><p style="margin:0; font-size:0.7rem; color:#065f46">${c.rubro}</p></div>
                                </div>
                            </a>`;
                        });
                    }

                    // Render Novedades
                    if (matchedNews.length > 0) {
                        res.innerHTML += `<div class="bot-results-title" style="font-size: 0.72rem; font-weight: 800; color: #f59e0b; text-transform: uppercase; letter-spacing: 0.5px; margin: 15px 0 10px 5px;">NOVEDADES RECIENTES</div>`;
                        matchedNews.slice(0, 3).forEach(n => {
                            res.innerHTML += `<a href="novedades.html" class="res-item-bot" style="display:flex; align-items:center; gap:12px; padding:12px; border-radius:15px; margin-bottom:8px; text-decoration:none; color:inherit; background:#fffbeb; border:1px solid #fef3c7; display:block">
                                <div style="display:flex; align-items:center; gap:10px">
                                    <div style="width:30px; height:30px; border-radius:8px; background:#f59e0b; color:white; display:flex; align-items:center; justify-content:center; font-size:0.8rem"><i class="fa-solid fa-newspaper"></i></div>
                                    <div><h4 style="margin:0; font-size:0.85rem">${n.titulo}</h4></div>
                                </div>
                            </a>`;
                        });
                    }

                }, 250); // Búsqueda ultra-rápida de 250ms simulando tiempo real
            });
        }
    });

})();
