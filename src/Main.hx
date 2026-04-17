import h2d.Text;
import h2d.Graphics;
import h2d.Interactive;
import hxd.Key;
import hxd.res.DefaultFont;

class Main extends hxd.App {
    // Config
    static var DOMAIN:String = "";
    static var DEFAULT_USER:String = "";
    static var SHARE_PATH:String = "";

    // State
    var screen:String = "login";
    var font:h2d.Font;
    var centerX:Float;
    var centerY:Float;
    var W:Float;
    var H:Float;

    // Login
    var usernameStr:String = "";
    var passwordStr:String = "";
    var activeField:String = "password";
    var usernameDisplay:Text;
    var passwordDisplay:Text;
    var messageText:Text;
    var usernameBorder:Graphics;
    var passwordBorder:Graphics;
    var userFieldY:Float = 0;
    var passFieldY:Float = 0;
    var fieldX:Float = 0;
    var fieldW:Float = 280;
    var fieldH:Float = 28;

    // Dashboard
    var gpoData:Array<Dynamic> = [];
    var dashLayer:h2d.Object;
    var statusText:Text;
    var scrollY:Float = 0;

    // Computers
    var computers:Array<Dynamic> = [];
    var selectedPC:Dynamic = null;
    var compLayer:h2d.Object;
    var compScrollY:Float = 0;
    var compSearchStr:String = "";
    var compSearchDisplay:Text;
    var compSearchBorder:Graphics;

    // App Blocking
    var appLayer:h2d.Object;
    var appScrollY:Float = 0;
    var blockedApps:Array<String> = [];
    var appCheckStates:Map<String, Bool>;
    var appSearchStr:String = "";
    var appSearchDisplay:Text;
    var appSearchBorder:Graphics;
    var appStatusText:Text;
    var editingSearch:Bool = false;

    // All known Windows apps (process names)
    static var ALL_WINDOWS_APPS:Array<Array<String>> = [
        // [processName, displayName, category]
        // ── Acessorios ──
        ["calc", "Calculadora (Classica)", "Acessorios"],
        ["calculatorapp", "Calculadora (UWP)", "Acessorios"],
        ["mspaint", "Paint (Classico)", "Acessorios"],
        ["paintapp", "Paint (UWP)", "Acessorios"],
        ["notepad", "Bloco de Notas (Classico)", "Acessorios"],
        ["notepadapp", "Bloco de Notas (UWP)", "Acessorios"],
        ["wordpad", "WordPad", "Acessorios"],
        ["snippingtool", "Ferramenta de Recorte", "Acessorios"],
        ["screensketch", "Recorte e Anotacao", "Acessorios"],
        ["soundrecorder", "Gravador de Voz", "Acessorios"],
        ["stikynot", "Notas Adesivas", "Acessorios"],
        ["charmap", "Mapa de Caracteres", "Acessorios"],
        ["magnify", "Lupa", "Acessorios"],
        ["narrator", "Narrador", "Acessorios"],
        ["osk", "Teclado Virtual", "Acessorios"],
        // ── Navegadores ──
        ["msedge", "Microsoft Edge", "Navegadores"],
        ["chrome", "Google Chrome", "Navegadores"],
        ["firefox", "Mozilla Firefox", "Navegadores"],
        ["opera", "Opera", "Navegadores"],
        ["brave", "Brave Browser", "Navegadores"],
        ["iexplore", "Internet Explorer", "Navegadores"],
        // ── Comunicacao ──
        ["teams", "Microsoft Teams (Classico)", "Comunicacao"],
        ["ms-teams", "Microsoft Teams (Novo)", "Comunicacao"],
        ["outlook", "Microsoft Outlook", "Comunicacao"],
        ["discord", "Discord", "Comunicacao"],
        ["telegram", "Telegram", "Comunicacao"],
        ["whatsapp", "WhatsApp Desktop", "Comunicacao"],
        ["slack", "Slack", "Comunicacao"],
        ["skype", "Skype", "Comunicacao"],
        ["zoom", "Zoom", "Comunicacao"],
        // ── Entretenimento / Jogos ──
        ["steam", "Steam", "Jogos"],
        ["steamwebhelper", "Steam WebHelper", "Jogos"],
        ["epicgameslauncher", "Epic Games Launcher", "Jogos"],
        ["gog", "GOG Galaxy", "Jogos"],
        ["battle.net", "Battle.net", "Jogos"],
        ["minecraft", "Minecraft", "Jogos"],
        ["solitaire", "Solitaire", "Jogos"],
        ["xboxapp", "Xbox App", "Jogos"],
        ["gamebar", "Xbox Game Bar", "Jogos"],
        ["gamingoverlay", "Game Overlay", "Jogos"],
        // ── Midia ──
        ["wmplayer", "Windows Media Player", "Midia"],
        ["vlc", "VLC Media Player", "Midia"],
        ["spotify", "Spotify", "Midia"],
        ["groove", "Groove Music", "Midia"],
        ["movies", "Filmes e TV", "Midia"],
        ["photos", "Fotos", "Midia"],
        ["camera", "Camera", "Midia"],
        // ── Office ──
        ["winword", "Microsoft Word", "Office"],
        ["excel", "Microsoft Excel", "Office"],
        ["powerpnt", "Microsoft PowerPoint", "Office"],
        ["msaccess", "Microsoft Access", "Office"],
        ["onenote", "OneNote", "Office"],
        ["mspub", "Microsoft Publisher", "Office"],
        // ── Sistema ──
        ["taskmgr", "Gerenciador de Tarefas", "Sistema"],
        ["cmd", "Prompt de Comando", "Sistema"],
        ["powershell", "Windows PowerShell", "Sistema"],
        ["pwsh", "PowerShell 7+", "Sistema"],
        ["regedit", "Editor de Registro", "Sistema"],
        ["mmc", "Console de Gerenciamento", "Sistema"],
        ["control", "Painel de Controle", "Sistema"],
        ["systemsettings", "Configuracoes", "Sistema"],
        ["devmgmt", "Gerenciador de Dispositivos", "Sistema"],
        ["diskmgmt", "Gerenciamento de Disco", "Sistema"],
        ["msconfig", "Configuracao do Sistema", "Sistema"],
        ["perfmon", "Monitor de Desempenho", "Sistema"],
        ["resmon", "Monitor de Recursos", "Sistema"],
        ["eventvwr", "Visualizador de Eventos", "Sistema"],
        ["compmgmt", "Gerenciamento do Computador", "Sistema"],
        ["lusrmgr", "Usuarios e Grupos Locais", "Sistema"],
        ["gpedit", "Editor de Politica de Grupo", "Sistema"],
        ["secpol", "Politica de Seguranca Local", "Sistema"],
        ["services", "Servicos", "Sistema"],
        ["msinfo32", "Informacoes do Sistema", "Sistema"],
        ["dxdiag", "Ferramenta de Diagnostico", "Sistema"],
        ["cleanmgr", "Limpeza de Disco", "Sistema"],
        ["dfrgui", "Desfragmentador", "Sistema"],
        ["wt", "Windows Terminal", "Sistema"],
        // ── Rede / Transferencia ──
        ["mstsc", "Conexao de Area de Trabalho Remota", "Rede"],
        ["msra", "Assistencia Remota", "Rede"],
        ["qbittorrent", "qBittorrent", "Rede"],
        ["utorrent", "uTorrent", "Rede"],
        ["bittorrent", "BitTorrent", "Rede"],
        ["filezilla", "FileZilla", "Rede"],
        ["putty", "PuTTY", "Rede"],
        ["winscp", "WinSCP", "Rede"],
        ["teamviewer", "TeamViewer", "Rede"],
        ["anydesk", "AnyDesk", "Rede"],
        // ── Loja / UWP ──
        ["winstore.app", "Microsoft Store", "Loja"],
        ["yourphone", "Seu Telefone", "Loja"],
        ["cortana", "Cortana", "Loja"],
        ["bingweather", "Clima", "Loja"],
        ["bingnews", "Noticias", "Loja"],
        ["bingsports", "Esportes", "Loja"],
        ["bingfinance", "Financas", "Loja"],
        ["people", "Pessoas", "Loja"],
        ["windowsmaps", "Mapas", "Loja"],
        ["windowsalarms", "Alarmes e Relogio", "Loja"],
        ["gethelp", "Obter Ajuda", "Loja"],
        ["feedback", "Hub de Feedback", "Loja"],
        ["gamingapp", "Xbox Gaming", "Loja"],
        ["clipchamp", "Clipchamp", "Loja"],
        ["windowsterminal", "Windows Terminal (Store)", "Loja"],
        ["todoapp", "Microsoft To Do", "Loja"],
        // ── Desenvolvimento ──
        ["code", "Visual Studio Code", "Dev"],
        ["devenv", "Visual Studio", "Dev"],
        ["idea64", "IntelliJ IDEA", "Dev"],
        ["pycharm64", "PyCharm", "Dev"],
        ["git", "Git", "Dev"],
        ["node", "Node.js", "Dev"],
        ["python", "Python", "Dev"],
        ["java", "Java", "Dev"],
        // ── Outros ──
        ["onedrive", "OneDrive", "Outros"],
        ["dropbox", "Dropbox", "Outros"],
        ["googledrive", "Google Drive", "Outros"],
        ["acrord32", "Adobe Reader", "Outros"],
        ["acrobat", "Adobe Acrobat", "Outros"],
        ["7zfm", "7-Zip", "Outros"],
        ["winrar", "WinRAR", "Outros"],
        ["ccleaner", "CCleaner", "Outros"],
        ["malwarebytes", "Malwarebytes", "Outros"]
    ];

    override function init() {
        engine.backgroundColor = 0xFF1E1E2E;
        font = DefaultFont.get();
        W = s2d.width;
        H = s2d.height;
        centerX = W / 2;
        centerY = H / 2;
        appCheckStates = new Map();

        try {
            var p = new sys.io.Process("powershell", ["-NoProfile", "-Command", "(Get-WmiObject Win32_ComputerSystem).Domain"]);
            var o = StringTools.trim(p.stdout.readAll().toString()); p.close();
            if (o.length > 0 && o != "WORKGROUP") DOMAIN = o;
        } catch(_:Dynamic) {}
        try {
            var p = new sys.io.Process("powershell", ["-NoProfile", "-Command", "$env:USERNAME"]);
            var o = StringTools.trim(p.stdout.readAll().toString()); p.close();
            if (o.length > 0) { DEFAULT_USER = o; usernameStr = o; }
        } catch(_:Dynamic) {}

        // Detectar share path do config
        try {
            var cfgPath = "service/sysmon-config.json";
            if (sys.FileSystem.exists(cfgPath)) {
                var content = sys.io.File.getContent(cfgPath);
                var cfg = haxe.Json.parse(content);
                var rp:String = Reflect.field(cfg, "RemoteBlockedAppsPath");
                if (rp != null && rp.length > 0) SHARE_PATH = rp;
            }
        } catch(_:Dynamic) {}

        showLogin();
    }

    // ═══════════════════════════════════════
    //  LOGIN
    // ═══════════════════════════════════════
    function showLogin() {
        screen = "login";
        s2d.removeChildren();
        activeField = "password";
        fieldX = centerX - 140;

        var title = makeText("Bloqueio de Apps - Painel TI", 2, 0x89B4FA);
        title.x = centerX - (title.textWidth * 2) / 2; title.y = 40;

        var sub = makeText(DOMAIN.length > 0 ? "Dominio: " + DOMAIN : "(dominio nao detectado)", 1.2, 0x6C7086);
        sub.x = centerX - (sub.textWidth * 1.2) / 2; sub.y = 80;

        var ul = makeText("Usuario:", 1.5, 0xCDD6F4); ul.x = fieldX; ul.y = 120;
        userFieldY = 142;
        usernameBorder = new Graphics(s2d);
        drawFieldBox(usernameBorder, fieldX, userFieldY, fieldW, fieldH, false);
        usernameDisplay = makeText(usernameStr, 1.5, 0xFFFFFF);
        usernameDisplay.x = fieldX + 5; usernameDisplay.y = userFieldY + 5;
        var ua = new Interactive(fieldW, fieldH, s2d); ua.x = fieldX; ua.y = userFieldY; ua.cursor = Button;
        ua.onClick = function(_) { selectField("username"); };

        var pl = makeText("Senha:", 1.5, 0xCDD6F4); pl.x = fieldX; pl.y = 182;
        passFieldY = 204;
        passwordBorder = new Graphics(s2d);
        drawFieldBox(passwordBorder, fieldX, passFieldY, fieldW, fieldH, true);
        passwordDisplay = makeText("", 1.5, 0xFFFFFF);
        passwordDisplay.x = fieldX + 5; passwordDisplay.y = passFieldY + 5;
        var pa = new Interactive(fieldW, fieldH, s2d); pa.x = fieldX; pa.y = passFieldY; pa.cursor = Button;
        pa.onClick = function(_) { selectField("password"); };

        // Share path
        var sl = makeText("Share (blocked-apps.json):", 1.2, 0x6C7086); sl.x = fieldX; sl.y = 248;
        var shareFieldY = 266.0;
        var shareBorder = new Graphics(s2d);
        drawFieldBox(shareBorder, fieldX, shareFieldY, fieldW, fieldH, false);
        var shareDisp = makeText(SHARE_PATH.length > 0 ? SHARE_PATH : "service/blocked-apps.json", 1.1, 0xA6ADC8);
        shareDisp.x = fieldX + 5; shareDisp.y = shareFieldY + 6;

        makeButton("Entrar", centerX - 55, 310, 110, 30, function() { doLogin(); });

        messageText = makeText("", 1.3, 0xF38BA8); messageText.y = 355;
    }

    function selectField(field:String) {
        activeField = field;
        drawFieldBox(usernameBorder, fieldX, userFieldY, fieldW, fieldH, field == "username");
        drawFieldBox(passwordBorder, fieldX, passFieldY, fieldW, fieldH, field == "password");
    }

    function doLogin() {
        messageText.textColor = 0xF9E2AF; messageText.text = "Autenticando...";
        messageText.x = centerX - (messageText.textWidth * 1.3) / 2;
        var result = runBridge("auth", {domain: DOMAIN, username: usernameStr, password: passwordStr});
        if (result == null) { showMsg("Erro: bridge nao executou", 0xF38BA8); return; }
        var status:String = Reflect.field(result, "status");
        if (status == "ok") {
            var data:Dynamic = Reflect.field(result, "data");
            if (Reflect.field(data, "authenticated") == true) {
                showComputerList();
            } else {
                showMsg("Credenciais invalidas!", 0xF38BA8);
                passwordStr = ""; updateFieldDisplay();
            }
        } else { showMsg(Reflect.field(result, "message"), 0xF38BA8); }
    }

    function showMsg(msg:String, color:Int) {
        if (msg == null) msg = "Erro desconhecido";
        messageText.textColor = color; messageText.text = msg;
        messageText.x = centerX - (messageText.textWidth * 1.3) / 2;
    }

    // ═══════════════════════════════════════
    //  LISTA DE COMPUTADORES
    // ═══════════════════════════════════════
    function showComputerList() {
        screen = "computers";
        s2d.removeChildren();
        compScrollY = 0;
        compSearchStr = "";
        editingSearch = false;

        drawHeader("Selecione o Computador");

        // Busca
        var searchLabel = makeText("Buscar:", 1.3, 0xCDD6F4);
        searchLabel.x = 10; searchLabel.y = 48;
        compSearchBorder = new Graphics(s2d);
        drawFieldBox(compSearchBorder, 70, 46, 200, 22, false);
        compSearchDisplay = makeText("", 1.2, 0xFFFFFF);
        compSearchDisplay.x = 75; compSearchDisplay.y = 48;
        var searchArea = new Interactive(200, 22, s2d);
        searchArea.x = 70; searchArea.y = 46; searchArea.cursor = Button;
        searchArea.onClick = function(_) {
            editingSearch = true;
            drawFieldBox(compSearchBorder, 70, 46, 200, 22, true);
        };

        statusText = makeText("Carregando computadores do AD...", 1.2, 0xF9E2AF);
        statusText.x = 10; statusText.y = 74;

        compLayer = new h2d.Object(s2d);
        compLayer.x = 10; compLayer.y = 95;

        loadComputers();
    }

    function loadComputers() {
        var result = runBridge("list-computers", {domain: DOMAIN});
        if (result == null) { statusText.text = "Erro ao conectar"; statusText.textColor = 0xF38BA8; return; }
        var status:String = Reflect.field(result, "status");
        if (status == "ok") {
            var data:Dynamic = Reflect.field(result, "data");
            if (data != null && Std.isOfType(data, Array)) {
                computers = cast data;
                statusText.text = computers.length + " computadores encontrados";
                statusText.textColor = 0xA6E3A1;
                renderComputerList();
            } else { statusText.text = "Nenhum computador"; statusText.textColor = 0xF9E2AF; }
        } else { statusText.text = Reflect.field(result, "message"); statusText.textColor = 0xF38BA8; }
    }

    function renderComputerList() {
        compLayer.removeChildren();
        var y:Float = 0;
        var w = W - 20;
        var filter = compSearchStr.toLowerCase();

        // Header
        var hdr = new Graphics(compLayer);
        hdr.beginFill(0x45475A); hdr.drawRect(0, y, w, 22); hdr.endFill();
        addLayerText(compLayer, "Nome", 10, y + 3, 1.2, 0xCDD6F4);
        addLayerText(compLayer, "SO", 180, y + 3, 1.2, 0xCDD6F4);
        addLayerText(compLayer, "Ultimo Logon", 430, y + 3, 1.2, 0xCDD6F4);
        addLayerText(compLayer, "OU", 580, y + 3, 1.2, 0xCDD6F4);
        y += 24;

        var count = 0;
        for (i in 0...computers.length) {
            var pc:Dynamic = computers[i];
            var name:String = Reflect.field(pc, "name");
            if (name == null) name = "?";
            if (filter.length > 0 && name.toLowerCase().indexOf(filter) == -1) continue;

            var color:Int = (count % 2 == 0) ? 0x313244 : 0x1E1E2E;
            var row = new Graphics(compLayer);
            row.beginFill(color); row.drawRect(0, y, w, 22); row.endFill();

            addLayerText(compLayer, name, 10, y + 3, 1.1, 0x89B4FA);
            var os:String = Reflect.field(pc, "os"); if (os == null) os = "";
            addLayerText(compLayer, os, 180, y + 3, 1.0, 0xA6ADC8);
            var ll:String = Reflect.field(pc, "lastLogon"); if (ll == null) ll = "";
            addLayerText(compLayer, ll, 430, y + 3, 1.0, 0x6C7086);
            var ou:String = Reflect.field(pc, "ou"); if (ou == null) ou = "";
            addLayerText(compLayer, ou, 580, y + 3, 1.0, 0x6C7086);

            // Click area
            var area = new Interactive(w, 22, compLayer);
            area.x = 0; area.y = y; area.cursor = Button;
            var pcRef = pc;
            area.onClick = function(_) { selectedPC = pcRef; showAppBlocking(); };
            area.onOver = function(_) { row.clear(); row.beginFill(0x45475A); row.drawRect(0, area.y, w, 22); row.endFill(); };
            var capturedY = y;
            var capturedColor = color;
            area.onOut = function(_) { row.clear(); row.beginFill(capturedColor); row.drawRect(0, capturedY, w, 22); row.endFill(); };

            y += 22;
            count++;
        }
    }

    // ═══════════════════════════════════════
    //  TELA DE BLOQUEIO DE APPS
    // ═══════════════════════════════════════
    function showAppBlocking() {
        screen = "apps";
        s2d.removeChildren();
        appScrollY = 0;
        appSearchStr = "";
        editingSearch = false;

        var pcName:String = Reflect.field(selectedPC, "name");
        drawHeader("Bloquear Apps - " + pcName);

        // Botao voltar
        makeButton("< Voltar", 10, 46, 70, 22, function() { showComputerList(); });

        // Busca
        var sl2 = makeText("Buscar:", 1.3, 0xCDD6F4); sl2.x = 90; sl2.y = 48;
        appSearchBorder = new Graphics(s2d);
        drawFieldBox(appSearchBorder, 150, 46, 180, 22, false);
        appSearchDisplay = makeText("", 1.2, 0xFFFFFF);
        appSearchDisplay.x = 155; appSearchDisplay.y = 48;
        var sa = new Interactive(180, 22, s2d); sa.x = 150; sa.y = 46; sa.cursor = Button;
        sa.onClick = function(_) { editingSearch = true; drawFieldBox(appSearchBorder, 150, 46, 180, 22, true); };

        // Scope buttons
        makeButton("Salvar Global", Std.int(W) - 290, 46, 120, 22, function() { saveApps("global"); });
        makeButton("Salvar p/ PC", Std.int(W) - 160, 46, 120, 22, function() { saveApps("machine"); });

        appStatusText = makeText("", 1.2, 0xA6E3A1);
        appStatusText.x = 340; appStatusText.y = 48;

        // === Segunda linha: gerenciamento do agente no PC remoto ===
        makeButton("Instalar Agente", 10, 76, 130, 22, function() { installRemoteAgent(pcName); });
        makeButton("Remover Agente", 150, 76, 130, 22, function() { uninstallRemoteAgent(pcName); });
        makeButton("Status Agente", 290, 76, 120, 22, function() { checkRemoteAgent(pcName); });
        makeButton("Gerar .bat p/ GPO", Std.int(W) - 170, 76, 160, 22, function() { generateGpoBat(); });

        appLayer = new h2d.Object(s2d);
        appLayer.x = 10; appLayer.y = 106;

        // Load current blocks
        loadCurrentBlocks(pcName);
        renderAppList();
    }

    function loadCurrentBlocks(pcName:String) {
        appCheckStates = new Map();
        // Reset all to unchecked
        for (app in ALL_WINDOWS_APPS) appCheckStates.set(app[0], false);

        var sharePath = SHARE_PATH.length > 0 ? SHARE_PATH : "service/blocked-apps.json";
        var result = runBridge("get-blocked-apps", {sharePath: sharePath, hostname: pcName});
        if (result != null) {
            var status:String = Reflect.field(result, "status");
            if (status == "ok") {
                var data:Dynamic = Reflect.field(result, "data");
                var allApps:Dynamic = Reflect.field(data, "allApps");
                if (allApps != null && Std.isOfType(allApps, Array)) {
                    var arr:Array<Dynamic> = cast allApps;
                    for (item in arr) {
                        var s:String = Std.string(item).toLowerCase();
                        appCheckStates.set(s, true);
                    }
                }
            }
        }
    }

    function renderAppList() {
        appLayer.removeChildren();
        var y:Float = 0;
        var w = W - 20;
        var filter = appSearchStr.toLowerCase();
        var lastCategory = "";

        // Header
        var hdr = new Graphics(appLayer);
        hdr.beginFill(0x45475A); hdr.drawRect(0, y, w, 22); hdr.endFill();
        addLayerText(appLayer, "[X]", 10, y + 3, 1.2, 0xCDD6F4);
        addLayerText(appLayer, "Aplicativo", 40, y + 3, 1.2, 0xCDD6F4);
        addLayerText(appLayer, "Processo", 300, y + 3, 1.2, 0xCDD6F4);
        addLayerText(appLayer, "Categoria", 480, y + 3, 1.2, 0xCDD6F4);
        y += 24;

        for (i in 0...ALL_WINDOWS_APPS.length) {
            var app = ALL_WINDOWS_APPS[i];
            var procName = app[0];
            var displayName = app[1];
            var category = app[2];

            // Filter
            if (filter.length > 0) {
                if (procName.toLowerCase().indexOf(filter) == -1 &&
                    displayName.toLowerCase().indexOf(filter) == -1 &&
                    category.toLowerCase().indexOf(filter) == -1) continue;
            }

            // Category separator
            if (category != lastCategory) {
                var catBg = new Graphics(appLayer);
                catBg.beginFill(0x181825); catBg.drawRect(0, y, w, 18); catBg.endFill();
                addLayerText(appLayer, "── " + category + " ──", 10, y + 2, 1.1, 0xCBA6F7);
                y += 20;
                lastCategory = category;
            }

            var isBlocked = appCheckStates.exists(procName) && appCheckStates.get(procName);
            var rowColor:Int = isBlocked ? 0x45283C : ((i % 2 == 0) ? 0x313244 : 0x1E1E2E);

            var row = new Graphics(appLayer);
            row.beginFill(rowColor); row.drawRect(0, y, w, 20); row.endFill();

            // Checkbox
            var chk = new Graphics(appLayer);
            chk.lineStyle(1, 0x6C7086);
            chk.drawRect(12, y + 3, 14, 14);
            if (isBlocked) {
                chk.beginFill(0xF38BA8);
                chk.drawRect(14, y + 5, 10, 10);
                chk.endFill();
            }

            addLayerText(appLayer, displayName, 40, y + 3, 1.0, isBlocked ? 0xF38BA8 : 0xCDD6F4);
            addLayerText(appLayer, procName, 300, y + 3, 1.0, 0x6C7086);
            addLayerText(appLayer, category, 480, y + 3, 1.0, 0x585B70);

            // Click
            var area2 = new Interactive(w, 20, appLayer);
            area2.x = 0; area2.y = y; area2.cursor = Button;
            var pn = procName;
            area2.onClick = function(_) {
                var cur = appCheckStates.exists(pn) && appCheckStates.get(pn);
                appCheckStates.set(pn, !cur);
                renderAppList();
            };

            y += 20;
        }
    }

    function saveApps(scope:String) {
        var pcName:String = Reflect.field(selectedPC, "name");
        var sharePath = SHARE_PATH.length > 0 ? SHARE_PATH : "service/blocked-apps.json";
        var apps:Array<String> = [];
        for (app in ALL_WINDOWS_APPS) {
            if (appCheckStates.exists(app[0]) && appCheckStates.get(app[0])) {
                apps.push(app[0]);
            }
        }
        var result = runBridge("save-blocked-apps", {
            sharePath: sharePath,
            hostname: pcName,
            apps: apps,
            scope: scope
        });
        if (result != null && Reflect.field(result, "status") == "ok") {
            var label = (scope == "global") ? "GLOBAL" : pcName;
            appStatusText.text = "Salvo (" + label + "): " + apps.length + " apps bloqueados";
            appStatusText.textColor = 0xA6E3A1;
        } else {
            appStatusText.text = "Erro ao salvar!";
            appStatusText.textColor = 0xF38BA8;
        }
    }

    // ═══════════════════════════════════════
    //  GERENCIAMENTO DE AGENTE REMOTO
    // ═══════════════════════════════════════
    function setStatus(msg:String, color:Int) {
        if (appStatusText != null) { appStatusText.text = msg; appStatusText.textColor = color; }
    }

    function installRemoteAgent(pcName:String) {
        setStatus("Instalando agente em " + pcName + "...", 0xF9E2AF);
        var sharePath = SHARE_PATH.length > 0 ? SHARE_PATH : "service/blocked-apps.json";
        var result = runBridge("install-remote-agent", {
            hostname: pcName,
            sharePath: sharePath,
            username: usernameStr,
            password: passwordStr,
            domain: DOMAIN
        });
        if (result != null && Reflect.field(result, "status") == "ok") {
            setStatus("Agente instalado em " + pcName, 0xA6E3A1);
        } else {
            var msg = result != null ? Std.string(Reflect.field(result, "message")) : "bridge falhou";
            setStatus("Erro: " + msg, 0xF38BA8);
        }
    }

    function uninstallRemoteAgent(pcName:String) {
        setStatus("Removendo agente em " + pcName + "...", 0xF9E2AF);
        var result = runBridge("uninstall-remote-agent", {
            hostname: pcName,
            username: usernameStr,
            password: passwordStr,
            domain: DOMAIN
        });
        if (result != null && Reflect.field(result, "status") == "ok") {
            setStatus("Agente removido de " + pcName, 0xA6E3A1);
        } else {
            var msg = result != null ? Std.string(Reflect.field(result, "message")) : "bridge falhou";
            setStatus("Erro: " + msg, 0xF38BA8);
        }
    }

    function checkRemoteAgent(pcName:String) {
        setStatus("Consultando agente em " + pcName + "...", 0xF9E2AF);
        var result = runBridge("check-remote-agent", {
            hostname: pcName,
            username: usernameStr,
            password: passwordStr,
            domain: DOMAIN
        });
        if (result != null && Reflect.field(result, "status") == "ok") {
            var data:Dynamic = Reflect.field(result, "data");
            var inst:Dynamic = Reflect.field(data, "installed");
            if (inst == true) {
                var st:String = Std.string(Reflect.field(data, "state"));
                var lr:String = Std.string(Reflect.field(data, "lastRun"));
                setStatus(pcName + ": " + st + " (ultima exec: " + lr + ")", 0xA6E3A1);
            } else {
                setStatus(pcName + ": NAO INSTALADO", 0xF9E2AF);
            }
        } else {
            var msg = result != null ? Std.string(Reflect.field(result, "message")) : "bridge falhou";
            setStatus("Erro: " + msg, 0xF38BA8);
        }
    }

    function generateGpoBat() {
        setStatus("Gerando .bat p/ GPO...", 0xF9E2AF);
        var cwd = Sys.getCwd();
        var outBat = cwd + "scripts/install-winsysmon-gpo.bat";
        var sharePath = SHARE_PATH.length > 0 ? SHARE_PATH : "\\\\SERVIDOR\\share\\blocked-apps.json";
        var scriptShare = "\\\\SERVIDOR\\NETLOGON\\winsysmon.ps1";
        var result = runBridge("gen-gpo-bat", {
            sharePath: sharePath,
            scriptShare: scriptShare,
            outBat: outBat
        });
        if (result != null && Reflect.field(result, "status") == "ok") {
            setStatus("Gerado: " + outBat + "  (edite os caminhos UNC)", 0xA6E3A1);
        } else {
            var msg = result != null ? Std.string(Reflect.field(result, "message")) : "bridge falhou";
            setStatus("Erro: " + msg, 0xF38BA8);
        }
    }

    // ═══════════════════════════════════════
    //  BRIDGE POWERSHELL
    // ═══════════════════════════════════════
    function runBridge(cmd:String, args:Dynamic):Dynamic {
        var logFile = "scripts/_debug.log";
        function log(s:String) {
            try {
                var f = sys.io.File.append(logFile, false);
                f.writeString("[" + Date.now().toString() + "] " + s + "\n");
                f.close();
            } catch(_:Dynamic) {}
        }
        try {
            var basePath = Sys.getCwd();
            var inputFile = basePath + "scripts/_cmd.json";
            var outputFile = basePath + "scripts/_result.json";
            var bridgeFile = basePath + "scripts/bridge.ps1";
            if (!sys.FileSystem.exists(bridgeFile)) {
                log("bridge.ps1 NOT FOUND at " + bridgeFile);
                return null;
            }
            sys.io.File.saveContent(inputFile, haxe.Json.stringify({cmd: cmd, args: args}));

            var stderr = "";
            var ec = -1;
            try {
                var proc = new sys.io.Process("powershell.exe", [
                    "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                    "-File", bridgeFile,
                    "-InputFile", inputFile,
                    "-OutputFile", outputFile
                ]);
                proc.stdout.readAll();
                stderr = proc.stderr.readAll().toString();
                ec = proc.exitCode();
                proc.close();
            } catch(pe:Dynamic) {
                log("Process EXCEPTION cmd=" + cmd + ": " + Std.string(pe));
                try { sys.FileSystem.deleteFile(inputFile); } catch(_:Dynamic) {}
                return null;
            }

            try { sys.FileSystem.deleteFile(inputFile); } catch(_:Dynamic) {}

            if (!sys.FileSystem.exists(outputFile)) {
                log("output file NOT created cmd=" + cmd + " exit=" + ec + " stderr=" + stderr);
                return null;
            }
            var content = sys.io.File.getContent(outputFile);
            try { sys.FileSystem.deleteFile(outputFile); } catch(_:Dynamic) {}
            // Strip UTF-8 BOM (EF BB BF) if present
            if (content.length >= 3 && content.charCodeAt(0) == 0xEF && content.charCodeAt(1) == 0xBB && content.charCodeAt(2) == 0xBF) {
                content = content.substr(3);
            }
            // Also strip BOM encoded as single 0xFEFF codepoint (if decoded as UTF-16)
            if (content.length > 0 && content.charCodeAt(0) == 0xFEFF) {
                content = content.substr(1);
            }
            try {
                return haxe.Json.parse(content);
            } catch(je:Dynamic) {
                log("JSON parse fail cmd=" + cmd + ": " + Std.string(je) + " | content=" + content);
                return null;
            }
        } catch(e:Dynamic) {
            log("OUTER EXCEPTION cmd=" + cmd + ": " + Std.string(e));
            return null;
        }
    }

    // ═══════════════════════════════════════
    //  INPUT
    // ═══════════════════════════════════════
    override function update(dt:Float) {
        if (screen == "login") {
            handleLoginInput();
        } else if (screen == "computers") {
            handleScroll(compLayer, compScrollY, 95);
            handleSearchInput(function(s) { compSearchStr = s; compSearchDisplay.text = s; renderComputerList(); }, compSearchStr);
        } else if (screen == "apps") {
            handleScroll(appLayer, appScrollY, 76);
            handleSearchInput(function(s) { appSearchStr = s; appSearchDisplay.text = s; renderAppList(); }, appSearchStr);
        }
    }

    function handleScroll(layer:h2d.Object, curScroll:Float, baseY:Float) {
        if (layer == null) return;
        var delta:Float = 0;
        if (Key.isPressed(Key.MOUSE_WHEEL_DOWN)) delta = -30;
        if (Key.isPressed(Key.MOUSE_WHEEL_UP)) delta = 30;
        if (delta != 0) {
            var ns = curScroll + delta;
            if (ns > 0) ns = 0;
            if (screen == "computers") compScrollY = ns;
            else appScrollY = ns;
            layer.y = baseY + ns;
        }
    }

    function handleSearchInput(callback:String->Void, current:String) {
        if (!editingSearch) return;
        for (k in 0...256) {
            if (!Key.isPressed(k)) continue;
            if (k == Key.ESCAPE || k == Key.ENTER) { editingSearch = false; return; }
            if (k == Key.BACKSPACE) {
                if (current.length > 0) callback(current.substr(0, current.length - 1));
                return;
            }
            if (k >= Key.A && k <= Key.Z) {
                callback(current + String.fromCharCode(k - Key.A + 97));
                return;
            }
            if (k >= Key.NUMBER_0 && k <= Key.NUMBER_9) {
                callback(current + String.fromCharCode(k - Key.NUMBER_0 + 48));
                return;
            }
        }
    }

    function handleLoginInput() {
        for (k in 0...256) {
            if (!Key.isPressed(k)) continue;
            if (k == Key.TAB) { selectField(activeField == "username" ? "password" : "username"); }
            else if (k == Key.BACKSPACE) {
                if (activeField == "username" && usernameStr.length > 0)
                    usernameStr = usernameStr.substr(0, usernameStr.length - 1);
                else if (activeField == "password" && passwordStr.length > 0)
                    passwordStr = passwordStr.substr(0, passwordStr.length - 1);
                updateFieldDisplay();
            } else if (k == Key.ENTER) { doLogin(); }
            else if (k >= Key.A && k <= Key.Z) {
                var c = k - Key.A + 97;
                if (Key.isDown(Key.LSHIFT) || Key.isDown(Key.RSHIFT)) c -= 32;
                appendChar(String.fromCharCode(c));
            } else if (k >= Key.NUMBER_0 && k <= Key.NUMBER_9) {
                var ch = String.fromCharCode(k - Key.NUMBER_0 + 48);
                if (Key.isDown(Key.LSHIFT) || Key.isDown(Key.RSHIFT)) {
                    var sp = [")", "!", "@", "#", "$", "%", "^", "&", "*", "("];
                    ch = sp[k - Key.NUMBER_0];
                }
                appendChar(ch);
            } else if (k == Key.NUMPAD_DOT || k == 190) { appendChar("."); }
            else if (k == 189 || k == Key.NUMPAD_SUB) {
                appendChar(Key.isDown(Key.LSHIFT) || Key.isDown(Key.RSHIFT) ? "_" : "-");
            } else if (k == 220) { appendChar("\\"); }
        }
    }

    function appendChar(c:String) {
        if (activeField == "username") usernameStr += c;
        else passwordStr += c;
        updateFieldDisplay();
    }

    function updateFieldDisplay() {
        if (usernameDisplay != null) usernameDisplay.text = usernameStr;
        if (passwordDisplay != null) {
            var m = ""; for (_ in 0...passwordStr.length) m += "*";
            passwordDisplay.text = m;
        }
    }

    // ═══════════════════════════════════════
    //  UI HELPERS
    // ═══════════════════════════════════════
    function drawHeader(title:String) {
        var hdr = new Graphics(s2d);
        hdr.beginFill(0x313244); hdr.drawRect(0, 0, W, 40); hdr.endFill();
        var ht = makeText(title, 1.5, 0x89B4FA); ht.x = 10; ht.y = 10;
        var logout = makeText("[Sair]", 1.3, 0xF38BA8);
        logout.x = W - 60; logout.y = 12;
        var la = new Interactive(50, 20, s2d); la.x = W - 60; la.y = 12; la.cursor = Button;
        la.onClick = function(_) { passwordStr = ""; showLogin(); };
    }

    function makeText(str:String, scale:Float, color:Int):Text {
        var t = new Text(font, s2d); t.text = str; t.setScale(scale); t.textColor = color; return t;
    }

    function addLayerText(layer:h2d.Object, str:String, x:Float, y:Float, scale:Float, color:Int) {
        var t = new Text(font, layer); t.text = str; t.setScale(scale); t.textColor = color; t.x = x; t.y = y;
    }

    function makeButton(label:String, x:Float, y:Float, w:Float, h:Float, onClick:Void->Void) {
        var bg = new Graphics(s2d);
        bg.beginFill(0x89B4FA); bg.drawRect(x, y, w, h); bg.endFill();
        var t = makeText(label, 1.3, 0x1E1E2E);
        t.x = x + w / 2 - (t.textWidth * 1.3) / 2; t.y = y + 3;
        var area = new Interactive(w, h, s2d); area.x = x; area.y = y; area.cursor = Button;
        area.onClick = function(_) { onClick(); };
        area.onOver = function(_) { bg.clear(); bg.beginFill(0xB4D0FA); bg.drawRect(x, y, w, h); bg.endFill(); };
        area.onOut = function(_) { bg.clear(); bg.beginFill(0x89B4FA); bg.drawRect(x, y, w, h); bg.endFill(); };
    }

    function drawFieldBox(g:Graphics, x:Float, y:Float, w:Float, h:Float, active:Bool) {
        g.clear(); g.beginFill(0x313244); g.drawRect(x, y, w, h); g.endFill();
        g.lineStyle(1, active ? 0x89B4FA : 0x45475A); g.drawRect(x, y, w, h);
    }

    static function main() { new Main(); }
}
