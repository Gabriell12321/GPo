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
    static var HOSTS_SHARE_PATH:String = "";
    static var POLICIES_SHARE_PATH:String = "";

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

    // Host / IP Blocking (sites e IPs)
    var hostLayer:h2d.Object;
    var hostScrollY:Float = 0;
    var blockedHostsList:Array<String> = [];
    var hostInputStr:String = "";
    var hostInputDisplay:Text;
    var hostInputBorder:Graphics;
    var editingHostInput:Bool = false;
    var hostStatusText:Text;

    // Politicas Extras (Widgets/Noticias)
    var policyLayer:h2d.Object;
    var policyWidgets:Bool = false;
    var policyStatusText:Text;
    var policyHasMachine:Bool = false;

    // Hyper-V Manager
    var hvHost:String = "";
    var hvLayer:h2d.Object;
    var hvAltUser:String = "";  // Credencial alternativa (conta local Hyper-V)
    var hvAltPass:String = "";
    var hvAltDomain:String = "";
    var hvScrollY:Float = 0;
    var hvVms:Array<Dynamic> = [];
    var hvStatusText:Text;
    var hvHostDisplay:Text;
    var hvHostBorder:Graphics;
    var hvSelectedVM:Dynamic = null;
    var hvSnapshots:Array<Dynamic> = [];
    var hvSwitches:Array<Dynamic> = [];
    var hvAutoRefreshTimer:Float = 0;
    var hvEditingHost:Bool = false;
    // Create-VM form
    var hvFormFields:Map<String, String>;
    var hvFormDisplays:Map<String, Text>;
    var hvFormBorders:Map<String, Graphics>;
    var hvFormActive:String = "";
    var hvFormStatus:Text;
    // Text input modal (snapshot name, VM name)
    var modalActive:Bool = false;
    var modalTitle:String = "";
    var modalValue:String = "";
    var modalDisplay:Text;
    var modalCallback:String->Void;

    // MikroTik
    var mtHost:String = "172.26.0.1";
    var mtHostDisplay:Text;
    var mtHostBorder:Graphics;
    var mtEditingHost:Bool = false;
    var mtWinboxPort:Int = 18291;
    var mtWebfigPort:Int = 18080;
    var mtTelnetPort:Int = 18023;
    var mtStatusText:Text;
    var mtResultsLayer:h2d.Object;
    var mtUser:String = "";
    var mtPass:String = "";
    var mtUserDisplay:Text;
    var mtPassDisplay:Text;
    var mtUserBorder:Graphics;
    var mtPassBorder:Graphics;
    var mtEditingUser:Bool = false;
    var mtEditingPass:Bool = false;

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
                var rh:String = Reflect.field(cfg, "RemoteBlockedHostsPath");
                if (rh != null && rh.length > 0) HOSTS_SHARE_PATH = rh;
                var rpol:String = Reflect.field(cfg, "RemoteBlockedPoliciesPath");
                if (rpol != null && rpol.length > 0) POLICIES_SHARE_PATH = rpol;
            }
        } catch(_:Dynamic) {}
        // Fallback: deriva de SHARE_PATH trocando blocked-apps.json -> blocked-hosts.json
        if (HOSTS_SHARE_PATH.length == 0 && SHARE_PATH.length > 0) {
            HOSTS_SHARE_PATH = StringTools.replace(SHARE_PATH, "blocked-apps.json", "blocked-hosts.json");
        }
        if (POLICIES_SHARE_PATH.length == 0 && SHARE_PATH.length > 0) {
            POLICIES_SHARE_PATH = StringTools.replace(SHARE_PATH, "blocked-apps.json", "blocked-policies.json");
        }

        // Listener global de texto (resolve teclas especiais como . , / - em qualquer layout)
        hxd.Window.getInstance().addEventTarget(onWindowEvent);

        showLogin();
    }

    function onWindowEvent(e:hxd.Event) {
        if (e.kind != hxd.Event.EventKind.ETextInput) return;
        var ch = String.fromCharCode(e.charCode);
        if (ch == null || ch.length == 0) return;
        // Roteia para o campo focado no momento
        if (editingHostInput) {
            hostInputStr += ch;
            if (hostInputDisplay != null) hostInputDisplay.text = hostInputStr;
            e.propagate = false;
        }
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

        // Botao Hyper-V
        makeButton("Hyper-V", Std.int(W) - 190, 10, 80, 22, function() { showHyperV(); });
        makeButton("MikroTik", Std.int(W) - 280, 10, 85, 22, function() { showMikroTik(); });

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
        makeButton("Salvar Global", Std.int(W) - 420, 46, 120, 22, function() { saveApps("global"); });
        makeButton("Salvar p/ PC",  Std.int(W) - 290, 46, 120, 22, function() { saveApps("machine"); });
        makeButton("Usar Global",   Std.int(W) - 160, 46, 120, 22, function() { saveApps("clear"); });

        appStatusText = makeText("", 1.2, 0xA6E3A1);
        appStatusText.x = 340; appStatusText.y = 48;

        // === Segunda linha: gerenciamento do agente no PC remoto ===
        makeButton("Instalar Agente", 10, 76, 130, 22, function() { installRemoteAgent(pcName); });
        makeButton("Remover Agente", 150, 76, 130, 22, function() { uninstallRemoteAgent(pcName); });
        makeButton("Status Agente", 290, 76, 120, 22, function() { checkRemoteAgent(pcName); });
        makeButton("Sites/IPs", 420, 76, 100, 22, function() { showHostBlocking(); });
        makeButton("Politicas", 530, 76, 100, 22, function() { showPolicyBlocking(); });
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
                if (allApps != null) {
                    if (Std.isOfType(allApps, Array)) {
                        var arr:Array<Dynamic> = cast allApps;
                        for (item in arr) {
                            var s:String = Std.string(item).toLowerCase();
                            appCheckStates.set(s, true);
                        }
                    } else {
                        var s:String = Std.string(allApps).toLowerCase();
                        if (s.length > 0) appCheckStates.set(s, true);
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
            var label = switch (scope) { case "global": "GLOBAL"; case "clear": "HERDANDO GLOBAL"; default: pcName; };
            if (scope == "clear") {
                appStatusText.text = "Override removido - " + pcName + " agora usa lista GLOBAL";
                // Recarrega para mostrar o que veio do global
                loadCurrentBlocks(pcName);
                renderAppList();
            } else {
                appStatusText.text = "Salvo (" + label + "): " + apps.length + " apps bloqueados";
            }
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
    //  BLOQUEIO DE SITES / IPs
    // ═══════════════════════════════════════
    function hostsSharePath():String {
        if (HOSTS_SHARE_PATH.length > 0) return HOSTS_SHARE_PATH;
        if (SHARE_PATH.length > 0) return StringTools.replace(SHARE_PATH, "blocked-apps.json", "blocked-hosts.json");
        return "service/blocked-hosts.json";
    }

    function showHostBlocking() {
        screen = "hosts";
        s2d.removeChildren();
        hostScrollY = 0;
        hostInputStr = "";
        editingHostInput = false;

        var pcName:String = Reflect.field(selectedPC, "name");
        drawHeader("Bloquear Sites/IPs - " + pcName);

        // Voltar
        makeButton("< Voltar", 10, 46, 70, 22, function() { showAppBlocking(); });

        // Campo de input novo host/IP
        var lb = makeText("Site ou IP:", 1.3, 0xCDD6F4); lb.x = 90; lb.y = 48;
        hostInputBorder = new Graphics(s2d);
        drawFieldBox(hostInputBorder, 185, 46, 260, 22, false);
        hostInputDisplay = makeText("", 1.2, 0xFFFFFF);
        hostInputDisplay.x = 190; hostInputDisplay.y = 48;
        var sa = new Interactive(260, 22, s2d); sa.x = 185; sa.y = 46; sa.cursor = Button;
        sa.onClick = function(_) { editingHostInput = true; drawFieldBox(hostInputBorder, 185, 46, 260, 22, true); };

        makeButton("+ Adicionar", 455, 46, 100, 22, function() { addHostEntry(); });

        // Scope buttons
        makeButton("Salvar Global", Std.int(W) - 420, 46, 120, 22, function() { saveHosts("global"); });
        makeButton("Salvar p/ PC",  Std.int(W) - 290, 46, 120, 22, function() { saveHosts("machine"); });
        makeButton("Usar Global",   Std.int(W) - 160, 46, 120, 22, function() { saveHosts("clear"); });

        // Linha 2: instrucoes
        var hint = makeText("Exemplos: facebook.com  |  *.tiktok.com  |  8.8.8.8  |  10.0.0.0/24", 1.0, 0x6C7086);
        hint.x = 10; hint.y = 76;

        hostStatusText = makeText("", 1.2, 0xA6E3A1);
        hostStatusText.x = 10; hostStatusText.y = 94;

        hostLayer = new h2d.Object(s2d);
        hostLayer.x = 10; hostLayer.y = 120;

        loadBlockedHosts(pcName);
        renderHostList();
    }

    function loadBlockedHosts(pcName:String) {
        blockedHostsList = [];
        var result = runBridge("get-blocked-hosts", {sharePath: hostsSharePath(), hostname: pcName});
        if (result != null && Reflect.field(result, "status") == "ok") {
            var data:Dynamic = Reflect.field(result, "data");
            var all:Dynamic = Reflect.field(data, "allHosts");
            if (all != null) {
                if (Std.isOfType(all, Array)) {
                    var arr:Array<Dynamic> = cast all;
                    for (item in arr) blockedHostsList.push(Std.string(item));
                } else {
                    // ConvertTo-Json colapsa array de 1 elemento em escalar
                    var s = Std.string(all);
                    if (s.length > 0) blockedHostsList.push(s);
                }
            }
        }
    }

    function renderHostList() {
        if (hostLayer == null) return;
        hostLayer.removeChildren();
        var y:Float = 0;
        var w = W - 20;

        // Header
        var hdr = new Graphics(hostLayer);
        hdr.beginFill(0x45475A); hdr.drawRect(0, y, w, 22); hdr.endFill();
        addLayerText(hostLayer, "Host / IP bloqueado", 10, y + 3, 1.2, 0xCDD6F4);
        addLayerText(hostLayer, "Tipo", Std.int(w) - 220, y + 3, 1.2, 0xCDD6F4);
        addLayerText(hostLayer, "Acao", Std.int(w) - 90, y + 3, 1.2, 0xCDD6F4);
        y += 24;

        if (blockedHostsList.length == 0) {
            addLayerText(hostLayer, "(nenhum site/IP bloqueado - adicione acima)", 10, y + 3, 1.1, 0x6C7086);
            return;
        }

        for (i in 0...blockedHostsList.length) {
            var entry = blockedHostsList[i];
            var isIp = isIpLike(entry);
            var tipo = isIp ? "IP (firewall)" : "Site (hosts)";
            var tipoColor = isIp ? 0xFAB387 : 0x89DCEB;
            var rowColor:Int = (i % 2 == 0) ? 0x313244 : 0x1E1E2E;

            var row = new Graphics(hostLayer);
            row.beginFill(rowColor); row.drawRect(0, y, w, 22); row.endFill();

            addLayerText(hostLayer, entry, 10, y + 3, 1.1, 0xF38BA8);
            addLayerText(hostLayer, tipo, Std.int(w) - 220, y + 3, 1.0, tipoColor);

            // Remove button
            var btn = new Graphics(hostLayer);
            btn.beginFill(0x585B70); btn.drawRect(Std.int(w) - 90, y + 2, 80, 18); btn.endFill();
            addLayerText(hostLayer, "Remover", Std.int(w) - 82, y + 4, 1.0, 0xF38BA8);
            var area = new Interactive(80, 18, hostLayer);
            area.x = Std.int(w) - 90; area.y = y + 2; area.cursor = Button;
            var idx = i;
            area.onClick = function(_) {
                blockedHostsList.splice(idx, 1);
                renderHostList();
            };

            y += 22;
        }
    }

    function isIpLike(s:String):Bool {
        var ereg = new EReg("^[0-9]+\\.[0-9]+\\.[0-9]+\\.[0-9]+(\\/[0-9]+)?$", "");
        return ereg.match(s);
    }

    function addHostEntry() {
        var s = StringTools.trim(hostInputStr.toLowerCase());
        if (s.length == 0) {
            if (hostStatusText != null) { hostStatusText.text = "Digite um site ou IP primeiro"; hostStatusText.textColor = 0xF9E2AF; }
            return;
        }
        // Evita duplicata
        for (e in blockedHostsList) if (e == s) {
            if (hostStatusText != null) { hostStatusText.text = "Ja existe na lista: " + s; hostStatusText.textColor = 0xF9E2AF; }
            return;
        }
        blockedHostsList.push(s);
        hostInputStr = "";
        if (hostInputDisplay != null) hostInputDisplay.text = "";
        if (hostStatusText != null) { hostStatusText.text = "Adicionado: " + s + "  (clique Salvar para aplicar)"; hostStatusText.textColor = 0xA6E3A1; }
        renderHostList();
    }

    function saveHosts(scope:String) {
        var pcName:String = Reflect.field(selectedPC, "name");
        var result = runBridge("save-blocked-hosts", {
            sharePath: hostsSharePath(),
            hostname: pcName,
            hosts: blockedHostsList,
            scope: scope
        });
        if (result != null && Reflect.field(result, "status") == "ok") {
            var label = switch (scope) { case "global": "GLOBAL"; case "clear": "HERDANDO GLOBAL"; default: pcName; };
            if (hostStatusText != null) {
                if (scope == "clear") {
                    hostStatusText.text = "Override removido - " + pcName + " agora usa lista GLOBAL";
                    loadBlockedHosts(pcName);
                    renderHostList();
                } else {
                    hostStatusText.text = "Salvo (" + label + "): " + blockedHostsList.length + " entradas (aplica em ate 60s nos PCs)";
                }
                hostStatusText.textColor = 0xA6E3A1;
            }
        } else {
            var msg = result != null ? Std.string(Reflect.field(result, "message")) : "bridge falhou";
            if (hostStatusText != null) { hostStatusText.text = "Erro ao salvar: " + msg; hostStatusText.textColor = 0xF38BA8; }
        }
    }

    function handleHostInputKeys() {
        if (!editingHostInput) return;
        // Apenas teclas de controle - caracteres vem via onWindowEvent (ETextInput)
        if (Key.isPressed(Key.ESCAPE)) {
            editingHostInput = false;
            drawFieldBox(hostInputBorder, 185, 46, 260, 22, false);
            return;
        }
        if (Key.isPressed(Key.ENTER)) {
            editingHostInput = false;
            drawFieldBox(hostInputBorder, 185, 46, 260, 22, false);
            addHostEntry();
            return;
        }
        if (Key.isPressed(Key.BACKSPACE)) {
            if (hostInputStr.length > 0) {
                hostInputStr = hostInputStr.substr(0, hostInputStr.length - 1);
                if (hostInputDisplay != null) hostInputDisplay.text = hostInputStr;
            }
            return;
        }
    }

    // ═══════════════════════════════════════
    //  POLITICAS EXTRAS (Widgets/Noticias Win10+11)
    // ═══════════════════════════════════════
    function policiesSharePath():String {
        if (POLICIES_SHARE_PATH.length > 0) return POLICIES_SHARE_PATH;
        if (SHARE_PATH.length > 0) return StringTools.replace(SHARE_PATH, "blocked-apps.json", "blocked-policies.json");
        return "service/blocked-policies.json";
    }

    function showPolicyBlocking() {
        screen = "policies";
        s2d.removeChildren();

        var pcName:String = Reflect.field(selectedPC, "name");
        drawHeader("Politicas Extras - " + pcName);

        makeButton("< Voltar", 10, 46, 70, 22, function() { showAppBlocking(); });

        makeButton("Salvar Global", Std.int(W) - 420, 46, 120, 22, function() { savePolicies("global"); });
        makeButton("Salvar p/ PC",  Std.int(W) - 290, 46, 120, 22, function() { savePolicies("machine"); });
        makeButton("Usar Global",   Std.int(W) - 160, 46, 120, 22, function() { savePolicies("clear"); });

        policyStatusText = makeText("", 1.2, 0xA6E3A1);
        policyStatusText.x = 10; policyStatusText.y = 78;

        policyLayer = new h2d.Object(s2d);
        policyLayer.x = 10; policyLayer.y = 110;

        loadBlockedPolicies(pcName);
        renderPolicyList();
    }

    function loadBlockedPolicies(pcName:String) {
        policyWidgets = false;
        policyHasMachine = false;
        var result = runBridge("get-blocked-policies", {sharePath: policiesSharePath(), hostname: pcName});
        if (result != null && Reflect.field(result, "status") == "ok") {
            var data:Dynamic = Reflect.field(result, "data");
            var eff:Dynamic = Reflect.field(data, "effective");
            if (eff != null) {
                var w:Dynamic = Reflect.field(eff, "Widgets");
                if (w != null) policyWidgets = (w == true || Std.string(w) == "True" || Std.string(w) == "true" || w == 1);
            }
            var hm:Dynamic = Reflect.field(data, "hasMachine");
            if (hm != null) policyHasMachine = (hm == true || Std.string(hm) == "True");
        }
    }

    function renderPolicyList() {
        if (policyLayer == null) return;
        policyLayer.removeChildren();
        var w = W - 20;

        // Header
        var hdr = new Graphics(policyLayer);
        hdr.beginFill(0x45475A); hdr.drawRect(0, 0, w, 22); hdr.endFill();
        addLayerText(policyLayer, "Politica / Ajuste", 10, 3, 1.2, 0xCDD6F4);
        addLayerText(policyLayer, "Escopo atual", Std.int(w) - 320, 3, 1.2, 0xCDD6F4);
        addLayerText(policyLayer, "Ativo?", Std.int(w) - 90, 3, 1.2, 0xCDD6F4);

        var y:Float = 30;

        // Widgets row
        var rowBg = new Graphics(policyLayer);
        rowBg.beginFill(0x313244); rowBg.drawRect(0, y, w, 56); rowBg.endFill();

        addLayerText(policyLayer, "Bloquear Widgets / Noticias do Windows", 10, y + 4, 1.3, 0xF38BA8);
        addLayerText(policyLayer, "- Remove o icone de Widgets (Win11) e News and Interests (Win10)", 10, y + 22, 1.0, 0xCDD6F4);
        addLayerText(policyLayer, "- Aplica Registry.pol + mata processos Widgets.exe / WidgetService.exe", 10, y + 36, 1.0, 0xCDD6F4);

        var scopeLabel = policyHasMachine ? "Override PC" : "Herda Global";
        var scopeColor = policyHasMachine ? 0xFAB387 : 0x89DCEB;
        addLayerText(policyLayer, scopeLabel, Std.int(w) - 320, y + 18, 1.1, scopeColor);

        // Toggle visual
        var boxX = Std.int(w) - 80;
        var boxY = Std.int(y) + 16;
        var box = new Graphics(policyLayer);
        var bgCol = policyWidgets ? 0xA6E3A1 : 0x585B70;
        box.beginFill(bgCol); box.drawRect(boxX, boxY, 60, 24); box.endFill();
        var knobX = policyWidgets ? boxX + 38 : boxX + 2;
        box.beginFill(0x1E1E2E); box.drawRect(knobX, boxY + 2, 20, 20); box.endFill();
        addLayerText(policyLayer, policyWidgets ? "SIM" : "NAO", boxX + (policyWidgets ? 6 : 30), boxY + 4, 1.0, policyWidgets ? 0x1E1E2E : 0xCDD6F4);

        var it = new Interactive(60, 24, policyLayer);
        it.x = boxX; it.y = boxY; it.cursor = Button;
        it.onClick = function(_) { policyWidgets = !policyWidgets; renderPolicyList(); };
    }

    function savePolicies(scope:String) {
        var pcName:String = Reflect.field(selectedPC, "name");
        var result = runBridge("save-blocked-policies", {
            sharePath: policiesSharePath(),
            hostname: pcName,
            widgets: policyWidgets,
            scope: scope
        });
        if (result != null && Reflect.field(result, "status") == "ok") {
            var label = switch (scope) { case "global": "GLOBAL"; case "clear": "HERDANDO GLOBAL"; default: pcName; };
            if (policyStatusText != null) {
                if (scope == "clear") {
                    policyStatusText.text = "Override removido - " + pcName + " agora usa lista GLOBAL";
                    loadBlockedPolicies(pcName);
                    renderPolicyList();
                } else {
                    policyStatusText.text = "Salvo (" + label + "): Widgets=" + (policyWidgets ? "BLOQUEADO" : "LIBERADO") + "  (aplica em ate 60s)";
                }
                policyStatusText.textColor = 0xA6E3A1;
            }
        } else {
            var msg = result != null ? Std.string(Reflect.field(result, "message")) : "bridge falhou";
            if (policyStatusText != null) { policyStatusText.text = "Erro: " + msg; policyStatusText.textColor = 0xF38BA8; }
        }
    }


    function setHvStatus(msg:String, color:Int) {
        if (hvStatusText != null) { hvStatusText.text = msg; hvStatusText.textColor = color; }
    }

    function hvUser():String { return hvAltUser.length > 0 ? hvAltUser : usernameStr; }
    function hvPass():String { return hvAltUser.length > 0 ? hvAltPass : passwordStr; }
    function hvDomain():String { return hvAltUser.length > 0 ? hvAltDomain : DOMAIN; }

    function openAltCredModal() {
        openTextModal("Usuario alt. (ex: .\\johnny.zack ou dom\\user) - ENTER", hvAltUser, function(u) {
            if (u == null || u.length == 0) {
                hvAltUser = ""; hvAltPass = ""; hvAltDomain = "";
                setHvStatus("Usando credenciais do login", 0xA6ADC8);
                return;
            }
            // Se comeca com .\ ou HOST\ deixamos em hvAltUser; dominio separado apenas se formato dom\user
            hvAltUser = u;
            hvAltDomain = "";
            if (u.indexOf("\\") > 0 && u.charAt(0) != ".") {
                var idx = u.indexOf("\\");
                hvAltDomain = u.substr(0, idx);
                hvAltUser = u.substr(idx + 1);
            } else if (u.charAt(0) == ".") {
                // .\user -> hvAltUser="user", sem dominio (bridge trata como local se user nao tem dominio)
                hvAltUser = u.substr(2);
            }
            openTextModal("Senha para " + u + " - ENTER", "", function(p) {
                hvAltPass = p;
                setHvStatus("Credencial alt. definida: " + u, 0xA6E3A1);
            });
        });
    }

    function showHyperV() {
        screen = "hyperv";
        s2d.removeChildren();
        hvScrollY = 0;
        hvAutoRefreshTimer = 0;
        hvEditingHost = false;
        editingSearch = false;
        modalActive = false;

        drawHeader("Hyper-V Manager");
        makeButton("< Voltar", 10, 46, 80, 22, function() { showComputerList(); });

        // Host input
        var lbl = makeText("Host:", 1.3, 0xCDD6F4); lbl.x = 100; lbl.y = 48;
        hvHostBorder = new Graphics(s2d);
        drawFieldBox(hvHostBorder, 145, 46, 200, 22, false);
        hvHostDisplay = makeText(hvHost, 1.2, 0xFFFFFF);
        hvHostDisplay.x = 150; hvHostDisplay.y = 48;
        var hArea = new Interactive(200, 22, s2d);
        hArea.x = 145; hArea.y = 46; hArea.cursor = Button;
        hArea.onClick = function(_) {
            hvEditingHost = true;
            drawFieldBox(hvHostBorder, 145, 46, 200, 22, true);
        };

        makeButton("Conectar", 355, 46, 90, 22, function() {
            hvEditingHost = false;
            drawFieldBox(hvHostBorder, 145, 46, 200, 22, false);
            loadVMs();
        });
        makeButton("Atualizar", 455, 46, 90, 22, function() { loadVMs(); });
        makeButton("Nova VM", 555, 46, 90, 22, function() {
            if (hvHost.length == 0) { setHvStatus("Informe o host primeiro", 0xF9E2AF); return; }
            showCreateVMForm();
        });
        makeButton("Testar", 655, 46, 70, 22, function() { testConnection(); });
        makeButton("Config WinRM", 735, 46, 110, 22, function() { configTrustedHosts(); });

        var credLbl = hvAltUser.length > 0 ? ("Cred: " + hvAltUser + " [mudar]") : "Usar cred alternativa";
        makeButton(credLbl, 855, 46, 250, 22, function() { openAltCredModal(); });

        hvStatusText = makeText(hvHost.length == 0 ? "Informe o host Hyper-V e clique Conectar" : "Pronto", 1.2, 0xF9E2AF);
        hvStatusText.x = 10; hvStatusText.y = 76;

        hvLayer = new h2d.Object(s2d);
        hvLayer.x = 10; hvLayer.y = 100;

        if (hvHost.length > 0 && hvVms.length > 0) {
            renderVMList();
        }
    }

    function testConnection() {
        if (hvHost.length == 0) { setHvStatus("Host vazio", 0xF38BA8); return; }
        setHvStatus("Testando " + hvHost + "...", 0xF9E2AF);
        var result = runBridge("hv-test-connection", {
            hvHost: hvHost, username: hvUser(), password: hvPass(), domain: hvDomain()
        });
        if (result == null) { setHvStatus("Bridge falhou", 0xF38BA8); return; }
        if (Reflect.field(result, "status") != "ok") {
            setHvStatus("Erro: " + Std.string(Reflect.field(result, "message")), 0xF38BA8);
            return;
        }
        var d:Dynamic = Reflect.field(result, "data");
        var p5985:String = Std.string(Reflect.field(d, "port5985"));
        var p5986:String = Std.string(Reflect.field(d, "port5986"));
        var authOk:Bool = Reflect.field(d, "authOk") == true;
        var hvOk:Bool = Reflect.field(d, "hvOk") == true;
        var method:String = Std.string(Reflect.field(d, "authMethod"));
        var err:String = Std.string(Reflect.field(d, "error"));
        var portStatus = "5985=" + p5985 + " 5986=" + p5986;
        if (!authOk) {
            setHvStatus(portStatus + " | Auth FALHOU: " + err, 0xF38BA8);
        } else if (!hvOk) {
            setHvStatus(portStatus + " | Auth OK (" + method + ") mas Hyper-V nao instalado", 0xF9E2AF);
        } else {
            setHvStatus(portStatus + " | Auth OK (" + method + ") | Hyper-V pronto", 0xA6E3A1);
        }
    }

    function configTrustedHosts() {
        if (hvHost.length == 0) { setHvStatus("Informe o host primeiro", 0xF9E2AF); return; }
        setHvStatus("Configurando TrustedHosts (sera solicitado UAC)...", 0xF9E2AF);
        var result = runBridge("hv-configure-trusted-hosts", { hosts: [hvHost] });
        if (result == null) { setHvStatus("Bridge falhou", 0xF38BA8); return; }
        if (Reflect.field(result, "status") == "ok") {
            var d:Dynamic = Reflect.field(result, "data");
            setHvStatus("WinRM configurado: TrustedHosts = " + Std.string(Reflect.field(d, "configured")), 0xA6E3A1);
        } else {
            setHvStatus("Erro: " + Std.string(Reflect.field(result, "message")), 0xF38BA8);
        }
    }

    function loadVMs() {
        if (hvHost.length == 0) { setHvStatus("Host vazio", 0xF38BA8); return; }
        setHvStatus("Listando VMs em " + hvHost + "...", 0xF9E2AF);
        var result = runBridge("hv-list-vms", {
            hvHost: hvHost, username: hvUser(), password: hvPass(), domain: hvDomain()
        });
        if (result == null) { setHvStatus("Bridge falhou", 0xF38BA8); return; }
        var status:String = Reflect.field(result, "status");
        if (status != "ok") {
            setHvStatus("Erro: " + Std.string(Reflect.field(result, "message")), 0xF38BA8);
            return;
        }
        var data:Dynamic = Reflect.field(result, "data");
        var vmsD:Dynamic = Reflect.field(data, "vms");
        if (vmsD != null && Std.isOfType(vmsD, Array)) {
            hvVms = cast vmsD;
            setHvStatus(hvVms.length + " VMs encontradas em " + hvHost + " (auto-refresh 3s)", 0xA6E3A1);
            renderVMList();
        } else {
            hvVms = [];
            setHvStatus("Nenhuma VM", 0xF9E2AF);
            renderVMList();
        }
    }

    function stateColor(state:String):Int {
        return switch (state) {
            case "Running": 0xA6E3A1;
            case "Off": 0x6C7086;
            case "Paused" | "Saved": 0xF9E2AF;
            case "Starting" | "Stopping" | "Saving" | "Pausing" | "Resuming": 0x89DCEB;
            default: 0xA6ADC8;
        }
    }

    function renderVMList() {
        if (hvLayer == null) return;
        hvLayer.removeChildren();
        var y:Float = 0;
        var w = W - 20;

        // Header
        var hdr = new Graphics(hvLayer);
        hdr.beginFill(0x45475A); hdr.drawRect(0, y, w, 22); hdr.endFill();
        addLayerText(hvLayer, "VM", 10, y + 3, 1.2, 0xCDD6F4);
        addLayerText(hvLayer, "Estado", 180, y + 3, 1.2, 0xCDD6F4);
        addLayerText(hvLayer, "CPU%", 270, y + 3, 1.2, 0xCDD6F4);
        addLayerText(hvLayer, "RAM (MB)", 320, y + 3, 1.2, 0xCDD6F4);
        addLayerText(hvLayer, "Uptime", 420, y + 3, 1.2, 0xCDD6F4);
        addLayerText(hvLayer, "Acoes", 520, y + 3, 1.2, 0xCDD6F4);
        y += 24;

        for (i in 0...hvVms.length) {
            var vm:Dynamic = hvVms[i];
            var name:String = Std.string(Reflect.field(vm, "name"));
            var state:String = Std.string(Reflect.field(vm, "state"));
            var cpu:Int = Reflect.field(vm, "cpuUsage");
            var memA:Int = Reflect.field(vm, "memAssigned");
            var memD:Int = Reflect.field(vm, "memDemand");
            var uptime:String = Std.string(Reflect.field(vm, "uptime"));

            var rowColor:Int = (i % 2 == 0) ? 0x313244 : 0x1E1E2E;
            var row = new Graphics(hvLayer);
            row.beginFill(rowColor); row.drawRect(0, y, w, 28); row.endFill();

            addLayerText(hvLayer, name, 10, y + 6, 1.1, 0x89B4FA);
            addLayerText(hvLayer, state, 180, y + 6, 1.0, stateColor(state));
            addLayerText(hvLayer, Std.string(cpu) + "%", 270, y + 6, 1.0, 0xCDD6F4);
            var memStr = Std.string(memA);
            if (memD > 0 && memD != memA) memStr += "/" + Std.string(memD);
            addLayerText(hvLayer, memStr, 320, y + 6, 1.0, 0xCDD6F4);
            addLayerText(hvLayer, uptime, 420, y + 6, 0.9, 0x6C7086);

            // Action buttons inline
            var bx:Float = 520;
            var isOff = (state == "Off" || state == "Saved" || state == "Paused");
            var vmRef = vm;
            var nRef = name;

            // Start / Resume
            if (isOff) {
                makeLayerButton(hvLayer, "Start", bx, y + 4, 55, 20, 0xA6E3A1, function() { vmAction(nRef, "start"); });
            } else {
                makeLayerButton(hvLayer, "Pause", bx, y + 4, 55, 20, 0xF9E2AF, function() { vmAction(nRef, "save"); });
            }
            bx += 60;
            makeLayerButton(hvLayer, "Stop", bx, y + 4, 50, 20, 0xF38BA8, function() { vmAction(nRef, "shutdown"); });
            bx += 55;
            makeLayerButton(hvLayer, "Off!", bx, y + 4, 45, 20, 0xEB6F92, function() { vmAction(nRef, "stop"); });
            bx += 50;
            makeLayerButton(hvLayer, "Snap", bx, y + 4, 50, 20, 0xCBA6F7, function() {
                hvSelectedVM = vmRef;
                openTextModal("Nome do snapshot (Enter p/ confirmar)", "", function(nm) {
                    createSnapshot(nRef, nm);
                });
            });
            bx += 55;
            makeLayerButton(hvLayer, "Detalhes", bx, y + 4, 75, 20, 0x89DCEB, function() {
                hvSelectedVM = vmRef; showVMDetails();
            });
            bx += 80;
            makeLayerButton(hvLayer, "Del", bx, y + 4, 40, 20, 0xF38BA8, function() {
                hvSelectedVM = vmRef;
                openTextModal("Digite DELETE para confirmar remocao de '" + nRef + "'", "", function(v) {
                    if (v == "DELETE") { deleteVM(nRef, false); }
                    else setHvStatus("Remocao cancelada", 0xF9E2AF);
                });
            });

            y += 30;
        }
    }

    function vmAction(vmName:String, act:String) {
        setHvStatus("Executando '" + act + "' em " + vmName + "...", 0xF9E2AF);
        var result = runBridge("hv-vm-action", {
            hvHost: hvHost, vmName: vmName, action: act,
            username: hvUser(), password: hvPass(), domain: hvDomain()
        });
        if (result != null && Reflect.field(result, "status") == "ok") {
            setHvStatus("OK: " + vmName + " -> " + act, 0xA6E3A1);
            hvAutoRefreshTimer = 1.5;
        } else {
            var msg = result != null ? Std.string(Reflect.field(result, "message")) : "bridge falhou";
            setHvStatus("Erro: " + msg, 0xF38BA8);
        }
    }

    function createSnapshot(vmName:String, snapName:String) {
        setHvStatus("Criando snapshot de " + vmName + "...", 0xF9E2AF);
        var result = runBridge("hv-snapshot-action", {
            hvHost: hvHost, vmName: vmName, action: "create", newName: snapName,
            username: hvUser(), password: hvPass(), domain: hvDomain()
        });
        if (result != null && Reflect.field(result, "status") == "ok") {
            var data:Dynamic = Reflect.field(result, "data");
            setHvStatus("Snapshot criado: " + Std.string(Reflect.field(data, "created")), 0xA6E3A1);
        } else {
            var msg = result != null ? Std.string(Reflect.field(result, "message")) : "bridge falhou";
            setHvStatus("Erro: " + msg, 0xF38BA8);
        }
    }

    function deleteVM(vmName:String, delVhd:Bool) {
        setHvStatus("Removendo VM " + vmName + "...", 0xF9E2AF);
        var result = runBridge("hv-delete-vm", {
            hvHost: hvHost, vmName: vmName, deleteVhd: delVhd,
            username: hvUser(), password: hvPass(), domain: hvDomain()
        });
        if (result != null && Reflect.field(result, "status") == "ok") {
            setHvStatus("VM removida: " + vmName, 0xA6E3A1);
            loadVMs();
        } else {
            var msg = result != null ? Std.string(Reflect.field(result, "message")) : "bridge falhou";
            setHvStatus("Erro: " + msg, 0xF38BA8);
        }
    }

    // ═════════════ DETALHES + SNAPSHOTS ═════════════
    function showVMDetails() {
        screen = "hyperv-details";
        s2d.removeChildren();
        hvAutoRefreshTimer = 0;
        modalActive = false;

        var name:String = Std.string(Reflect.field(hvSelectedVM, "name"));
        drawHeader("VM: " + name);
        makeButton("< Voltar", 10, 46, 80, 22, function() { showHyperV(); });
        makeButton("Atualizar", 100, 46, 90, 22, function() { showVMDetails(); });
        makeButton("Novo Snapshot", 200, 46, 120, 22, function() {
            openTextModal("Nome do snapshot", "", function(nm) {
                createSnapshot(name, nm);
                haxe.Timer.delay(function() { showVMDetails(); }, 800);
            });
        });

        hvStatusText = makeText("Carregando stats...", 1.2, 0xF9E2AF);
        hvStatusText.x = 10; hvStatusText.y = 76;

        // Stats box
        var statsResult = runBridge("hv-vm-stats", {
            hvHost: hvHost, vmName: name,
            username: hvUser(), password: hvPass(), domain: hvDomain()
        });
        if (statsResult != null && Reflect.field(statsResult, "status") == "ok") {
            var s:Dynamic = Reflect.field(statsResult, "data");
            var y:Float = 100;
            var box = new Graphics(s2d);
            box.beginFill(0x313244); box.drawRect(10, y, W - 20, 150); box.endFill();
            var col1x:Float = 20; var col2x:Float = 250; var col3x:Float = 500;
            var ry = y + 10;
            var state:String = Std.string(Reflect.field(s, "state"));
            addFieldLine(col1x, ry, "Estado:", state, stateColor(state));
            addFieldLine(col1x, ry + 22, "CPU:", Std.string(Reflect.field(s, "cpuUsage")) + "%", 0xF9E2AF);
            addFieldLine(col1x, ry + 44, "Processadores:", Std.string(Reflect.field(s, "processors")), 0xCDD6F4);
            addFieldLine(col1x, ry + 66, "Geracao:", Std.string(Reflect.field(s, "generation")), 0xCDD6F4);
            addFieldLine(col1x, ry + 88, "Uptime:", Std.string(Reflect.field(s, "uptime")), 0x89DCEB);

            addFieldLine(col2x, ry, "RAM Atribuida:", Std.string(Reflect.field(s, "memAssigned")) + " MB", 0xA6E3A1);
            addFieldLine(col2x, ry + 22, "RAM Demanda:", Std.string(Reflect.field(s, "memDemand")) + " MB", 0xA6E3A1);
            addFieldLine(col2x, ry + 44, "RAM Inicial:", Std.string(Reflect.field(s, "memStartup")) + " MB", 0xCDD6F4);
            addFieldLine(col2x, ry + 66, "RAM Min:", Std.string(Reflect.field(s, "memMinimum")) + " MB", 0xCDD6F4);
            addFieldLine(col2x, ry + 88, "RAM Max:", Std.string(Reflect.field(s, "memMaximum")) + " MB", 0xCDD6F4);

            addFieldLine(col3x, ry, "VHD Tam:", Std.string(Reflect.field(s, "vhdSizeGB")) + " GB", 0xCBA6F7);
            addFieldLine(col3x, ry + 22, "VHD Usado:", Std.string(Reflect.field(s, "vhdUsedGB")) + " GB", 0xCBA6F7);
            addFieldLine(col3x, ry + 44, "Switch:", Std.string(Reflect.field(s, "switchName")), 0x89B4FA);
            addFieldLine(col3x, ry + 66, "MAC:", Std.string(Reflect.field(s, "macAddress")), 0x6C7086);
            addFieldLine(col3x, ry + 88, "IPs:", Std.string(Reflect.field(s, "ipAddresses")), 0x6C7086);

            setHvStatus("Stats atualizados", 0xA6E3A1);
        } else {
            setHvStatus("Erro ao obter stats", 0xF38BA8);
        }

        // Snapshots list
        var snapResult = runBridge("hv-list-snapshots", {
            hvHost: hvHost, vmName: name,
            username: hvUser(), password: hvPass(), domain: hvDomain()
        });
        var y2:Float = 260;
        addLayerText(s2d, "═══ Snapshots ═══", 10, y2, 1.3, 0xCBA6F7);
        y2 += 25;

        if (snapResult != null && Reflect.field(snapResult, "status") == "ok") {
            var data:Dynamic = Reflect.field(snapResult, "data");
            var snapsD:Dynamic = Reflect.field(data, "snapshots");
            if (snapsD != null && Std.isOfType(snapsD, Array)) {
                hvSnapshots = cast snapsD;
                if (hvSnapshots.length == 0) {
                    addLayerText(s2d, "Sem snapshots", 10, y2, 1.1, 0x6C7086);
                } else {
                    var hdr = new Graphics(s2d);
                    hdr.beginFill(0x45475A); hdr.drawRect(10, y2, W - 20, 22); hdr.endFill();
                    addLayerText(s2d, "Nome", 20, y2 + 3, 1.1, 0xCDD6F4);
                    addLayerText(s2d, "Criado", 280, y2 + 3, 1.1, 0xCDD6F4);
                    addLayerText(s2d, "Parent", 450, y2 + 3, 1.1, 0xCDD6F4);
                    addLayerText(s2d, "Acoes", W - 200, y2 + 3, 1.1, 0xCDD6F4);
                    y2 += 24;
                    for (i in 0...hvSnapshots.length) {
                        var sn:Dynamic = hvSnapshots[i];
                        var snm:String = Std.string(Reflect.field(sn, "name"));
                        var rc:Int = (i % 2 == 0) ? 0x313244 : 0x1E1E2E;
                        var r = new Graphics(s2d);
                        r.beginFill(rc); r.drawRect(10, y2, W - 20, 22); r.endFill();
                        addLayerText(s2d, snm, 20, y2 + 3, 1.0, 0x89B4FA);
                        addLayerText(s2d, Std.string(Reflect.field(sn, "created")), 280, y2 + 3, 0.95, 0xA6ADC8);
                        addLayerText(s2d, Std.string(Reflect.field(sn, "parent")), 450, y2 + 3, 0.95, 0x6C7086);
                        var snRef = snm;
                        makeButton("Restaurar", W - 195, y2 + 1, 85, 20, function() {
                            snapshotAction(name, snRef, "restore");
                        });
                        makeButton("Remover", W - 105, y2 + 1, 80, 20, function() {
                            snapshotAction(name, snRef, "delete");
                        });
                        y2 += 22;
                    }
                }
            }
        } else {
            addLayerText(s2d, "Erro ao listar snapshots", 10, y2, 1.1, 0xF38BA8);
        }
    }

    function addFieldLine(x:Float, y:Float, lbl:String, val:String, color:Int) {
        var t1 = makeText(lbl, 1.1, 0xA6ADC8); t1.x = x; t1.y = y;
        var t2 = makeText(val, 1.1, color); t2.x = x + 115; t2.y = y;
    }

    function snapshotAction(vmName:String, snapName:String, act:String) {
        setHvStatus(act + " " + snapName + "...", 0xF9E2AF);
        var result = runBridge("hv-snapshot-action", {
            hvHost: hvHost, vmName: vmName, snapshotName: snapName, action: act,
            username: hvUser(), password: hvPass(), domain: hvDomain()
        });
        if (result != null && Reflect.field(result, "status") == "ok") {
            setHvStatus("OK: " + act + " " + snapName, 0xA6E3A1);
            haxe.Timer.delay(function() { showVMDetails(); }, 500);
        } else {
            var msg = result != null ? Std.string(Reflect.field(result, "message")) : "bridge falhou";
            setHvStatus("Erro: " + msg, 0xF38BA8);
        }
    }

    // ═════════════ CRIAR VM ═════════════
    function showCreateVMForm() {
        screen = "hyperv-create";
        s2d.removeChildren();
        modalActive = false;

        // Inicializar apenas se campos ainda nao existem (preservar edicao)
        if (hvFormFields == null) {
            hvFormFields = new Map();
            hvFormFields.set("name", "NovaVM");
            hvFormFields.set("memoryMB", "2048");
            hvFormFields.set("vhdSizeGB", "40");
            hvFormFields.set("processors", "2");
            hvFormFields.set("generation", "2");
            hvFormFields.set("switchName", "");
            hvFormActive = "name";
        }
        hvFormDisplays = new Map();
        hvFormBorders = new Map();

        drawHeader("Criar VM em " + hvHost);
        makeButton("< Voltar", 10, 46, 80, 22, function() {
            hvFormFields = null; // reset form ao sair
            showHyperV();
        });

        // Carregar switches apenas se ainda nao carregou
        var firstSw = "";
        if (hvSwitches == null || hvSwitches.length == 0) {
            var swRes = runBridge("hv-list-switches", {
                hvHost: hvHost, username: hvUser(), password: hvPass(), domain: hvDomain()
            });
            if (swRes != null && Reflect.field(swRes, "status") == "ok") {
                var d:Dynamic = Reflect.field(swRes, "data");
                var swD:Dynamic = Reflect.field(d, "switches");
                if (swD != null && Std.isOfType(swD, Array)) {
                    hvSwitches = cast swD;
                }
            }
        }
        if (hvSwitches != null && hvSwitches.length > 0) {
            firstSw = Std.string(Reflect.field(hvSwitches[0], "name"));
            if (hvFormFields.get("switchName") == "") hvFormFields.set("switchName", firstSw);
        }

        var startY:Float = 90;
        var lh:Float = 42;
        var fields:Array<Array<String>> = [
            ["name", "Nome da VM"],
            ["memoryMB", "Memoria RAM (MB)"],
            ["vhdSizeGB", "Disco VHDX (GB)"],
            ["processors", "Processadores"],
            ["generation", "Geracao (1 ou 2)"],
            ["switchName", "Virtual Switch" + (hvSwitches != null && hvSwitches.length > 0 ? " (sugerido: " + firstSw + ")" : "")]
        ];

        for (i in 0...fields.length) {
            var fKey = fields[i][0];
            var fLbl = fields[i][1];
            var fy = startY + i * lh;
            var lblT = makeText(fLbl + ":", 1.2, 0xCDD6F4); lblT.x = 20; lblT.y = fy + 4;
            var bord = new Graphics(s2d);
            drawFieldBox(bord, 240, fy, 300, 24, fKey == hvFormActive);
            var disp = makeText(hvFormFields.get(fKey), 1.2, 0xFFFFFF);
            disp.x = 245; disp.y = fy + 4;
            hvFormDisplays.set(fKey, disp);
            hvFormBorders.set(fKey, bord);
            var ia = new Interactive(300, 24, s2d);
            ia.x = 240; ia.y = fy; ia.cursor = Button;
            var kRef = fKey;
            ia.onClick = function(_) { setFormActive(kRef); };
        }

        // Lista de switches disponiveis
        if (hvSwitches != null && hvSwitches.length > 0) {
            var sy = startY + fields.length * lh + 10;
            addLayerText(s2d, "Switches disponiveis:", 20, sy, 1.1, 0xA6ADC8);
            var names:Array<String> = [];
            for (sw in hvSwitches) names.push(Std.string(Reflect.field(sw, "name")));
            addLayerText(s2d, names.join(", "), 180, sy, 1.0, 0x6C7086);
        }

        makeButton("Criar VM", 240, startY + fields.length * lh + 40, 150, 30, function() { submitCreateVM(); });
        makeButton("Cancelar", 400, startY + fields.length * lh + 40, 140, 30, function() { showHyperV(); });

        hvFormStatus = makeText("", 1.2, 0xA6E3A1);
        hvFormStatus.x = 20; hvFormStatus.y = startY + fields.length * lh + 85;
    }

    function setFormActive(k:String) {
        hvFormActive = k;
        // Rebuild para atualizar highlight do campo
        showCreateVMForm();
    }

    function submitCreateVM() {
        var nm = hvFormFields.get("name");
        if (nm == null || nm.length == 0) { hvFormStatus.text = "Nome obrigatorio"; hvFormStatus.textColor = 0xF38BA8; return; }
        var memMB = Std.parseInt(hvFormFields.get("memoryMB"));
        var sizeGB = Std.parseInt(hvFormFields.get("vhdSizeGB"));
        var procs = Std.parseInt(hvFormFields.get("processors"));
        var gen = Std.parseInt(hvFormFields.get("generation"));
        if (memMB == null || memMB < 256) { hvFormStatus.text = "RAM minima 256 MB"; hvFormStatus.textColor = 0xF38BA8; return; }
        if (sizeGB == null || sizeGB < 1) { hvFormStatus.text = "Disco minimo 1 GB"; hvFormStatus.textColor = 0xF38BA8; return; }
        if (gen != 1 && gen != 2) { hvFormStatus.text = "Geracao deve ser 1 ou 2"; hvFormStatus.textColor = 0xF38BA8; return; }

        hvFormStatus.text = "Criando VM " + nm + "...";
        hvFormStatus.textColor = 0xF9E2AF;

        var result = runBridge("hv-create-vm", {
            hvHost: hvHost, vmName: nm, memoryMB: memMB, vhdSizeGB: sizeGB,
            processors: procs, generation: gen, switchName: hvFormFields.get("switchName"),
            username: hvUser(), password: hvPass(), domain: hvDomain()
        });
        if (result != null && Reflect.field(result, "status") == "ok") {
            hvFormStatus.text = "VM criada: " + nm + ". Voltando...";
            hvFormStatus.textColor = 0xA6E3A1;
            haxe.Timer.delay(function() { loadVMs(); showHyperV(); }, 1500);
        } else {
            var msg = result != null ? Std.string(Reflect.field(result, "message")) : "bridge falhou";
            hvFormStatus.text = "Erro: " + msg;
            hvFormStatus.textColor = 0xF38BA8;
        }
    }

    // ═════════════ MODAL DE TEXTO ═════════════
    function openTextModal(title:String, initial:String, cb:String->Void) {
        modalActive = true;
        modalTitle = title;
        modalValue = initial;
        modalCallback = cb;

        // Draw overlay
        var overlay = new Graphics(s2d);
        overlay.beginFill(0x000000, 0.7); overlay.drawRect(0, 0, W, H); overlay.endFill();
        overlay.name = "modalOverlay";

        var mw:Float = 500; var mh:Float = 140;
        var mx:Float = (W - mw) / 2; var my:Float = (H - mh) / 2;
        var box = new Graphics(s2d);
        box.beginFill(0x313244); box.drawRect(mx, my, mw, mh); box.endFill();
        box.lineStyle(2, 0x89B4FA);
        box.drawRect(mx, my, mw, mh);
        box.name = "modalOverlay";

        var ttl = makeText(title, 1.2, 0xCDD6F4); ttl.x = mx + 15; ttl.y = my + 12;
        ttl.name = "modalOverlay";

        var bord = new Graphics(s2d);
        drawFieldBox(bord, mx + 15, my + 55, mw - 30, 28, true);
        bord.name = "modalOverlay";

        modalDisplay = makeText(initial, 1.3, 0xFFFFFF);
        modalDisplay.x = mx + 20; modalDisplay.y = my + 58;
        modalDisplay.name = "modalOverlay";

        var hint = makeText("ENTER=OK   ESC=Cancelar", 1.0, 0x6C7086);
        hint.x = mx + 15; hint.y = my + 110;
        hint.name = "modalOverlay";
    }

    function closeModal() {
        modalActive = false;
        // remove overlay children by name
        var i = s2d.numChildren - 1;
        while (i >= 0) {
            var c = s2d.getChildAt(i);
            if (c.name == "modalOverlay") s2d.removeChild(c);
            i--;
        }
    }

    function makeLayerButton(layer:h2d.Object, label:String, x:Float, y:Float, w:Float, h:Float, color:Int, onClick:Void->Void) {
        var bg = new Graphics(layer);
        bg.beginFill(color); bg.drawRect(x, y, w, h); bg.endFill();
        var t = new Text(font, layer); t.text = label; t.setScale(0.95); t.textColor = 0x1E1E2E;
        t.x = x + (w - t.textWidth * 0.95) / 2; t.y = y + 3;
        var ia = new Interactive(w, h, layer);
        ia.x = x; ia.y = y; ia.cursor = Button;
        ia.onClick = function(_) { onClick(); };
        ia.onOver = function(_) { bg.clear(); bg.beginFill(0xF5C2E7); bg.drawRect(x, y, w, h); bg.endFill(); };
        ia.onOut = function(_) { bg.clear(); bg.beginFill(color); bg.drawRect(x, y, w, h); bg.endFill(); };
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
        if (modalActive) {
            handleModalInput();
            return;
        }
        if (screen == "login") {
            handleLoginInput();
        } else if (screen == "computers") {
            handleScroll(compLayer, compScrollY, 95);
            handleSearchInput(function(s) { compSearchStr = s; compSearchDisplay.text = s; renderComputerList(); }, compSearchStr);
        } else if (screen == "apps") {
            handleScroll(appLayer, appScrollY, 76);
            handleSearchInput(function(s) { appSearchStr = s; appSearchDisplay.text = s; renderAppList(); }, appSearchStr);
        } else if (screen == "hosts") {
            handleScroll(hostLayer, hostScrollY, 120);
            handleHostInputKeys();
            if (!editingHostInput && Key.isPressed(Key.ESCAPE)) showAppBlocking();
        } else if (screen == "hyperv") {
            handleScroll(hvLayer, hvScrollY, 100);
            if (hvEditingHost) handleHostInput();
            // Auto-refresh a cada 3s
            if (hvHost.length > 0 && hvVms.length > 0) {
                hvAutoRefreshTimer += dt;
                if (hvAutoRefreshTimer >= 3.0) {
                    hvAutoRefreshTimer = 0;
                    silentRefreshVMs();
                }
            }
        } else if (screen == "hyperv-create") {
            handleFormInput();
        } else if (screen == "hyperv-details") {
            if (Key.isPressed(Key.ESCAPE)) showHyperV();
        } else if (screen == "mikrotik") {
            if (mtEditingHost) handleMtHostInput();
            else if (mtEditingUser) handleMtUserInput();
            else if (mtEditingPass) handleMtPassInput();
            else if (Key.isPressed(Key.ESCAPE)) showComputerList();
        }
    }

    function silentRefreshVMs() {
        if (hvHost.length == 0) return;
        var result = runBridge("hv-list-vms", {
            hvHost: hvHost, username: hvUser(), password: hvPass(), domain: hvDomain()
        });
        if (result != null && Reflect.field(result, "status") == "ok") {
            var data:Dynamic = Reflect.field(result, "data");
            var vmsD:Dynamic = Reflect.field(data, "vms");
            if (vmsD != null && Std.isOfType(vmsD, Array)) {
                hvVms = cast vmsD;
                renderVMList();
            }
        }
    }

    function handleHostInput() {
        for (k in 0...256) {
            if (!Key.isPressed(k)) continue;
            if (k == Key.ENTER) {
                hvEditingHost = false;
                drawFieldBox(hvHostBorder, 145, 46, 200, 22, false);
                loadVMs();
                return;
            }
            if (k == Key.ESCAPE) { hvEditingHost = false; drawFieldBox(hvHostBorder, 145, 46, 200, 22, false); return; }
            if (k == Key.BACKSPACE) {
                if (hvHost.length > 0) { hvHost = hvHost.substr(0, hvHost.length - 1); hvHostDisplay.text = hvHost; }
                return;
            }
            if (k >= Key.A && k <= Key.Z) {
                var c = k - Key.A + 97;
                if (Key.isDown(Key.LSHIFT) || Key.isDown(Key.RSHIFT)) c -= 32;
                hvHost += String.fromCharCode(c); hvHostDisplay.text = hvHost; return;
            }
            if (k >= Key.NUMBER_0 && k <= Key.NUMBER_9) {
                hvHost += String.fromCharCode(k - Key.NUMBER_0 + 48); hvHostDisplay.text = hvHost; return;
            }
            if (k == 189 || k == Key.NUMPAD_SUB) { hvHost += "-"; hvHostDisplay.text = hvHost; return; }
            if (k == 190 || k == Key.NUMPAD_DOT) { hvHost += "."; hvHostDisplay.text = hvHost; return; }
        }
    }

    function handleFormInput() {
        if (hvFormActive == null || hvFormActive == "") return;
        var cur = hvFormFields.get(hvFormActive);
        if (cur == null) cur = "";
        var disp = hvFormDisplays.get(hvFormActive);
        for (k in 0...256) {
            if (!Key.isPressed(k)) continue;
            if (k == Key.TAB) {
                // ciclar campo
                var keys = ["name","memoryMB","vhdSizeGB","processors","generation","switchName"];
                var idx = 0; for (i in 0...keys.length) if (keys[i] == hvFormActive) { idx = i; break; }
                var nextIdx = (idx + 1) % keys.length;
                setFormActive(keys[nextIdx]);
                return;
            }
            if (k == Key.ENTER) { submitCreateVM(); return; }
            if (k == Key.ESCAPE) { hvFormFields = null; showHyperV(); return; }
            if (k == Key.BACKSPACE) {
                if (cur.length > 0) { cur = cur.substr(0, cur.length - 1); hvFormFields.set(hvFormActive, cur); if (disp != null) disp.text = cur; }
                return;
            }
            var ch:String = null;
            if (k >= Key.A && k <= Key.Z) {
                var c = k - Key.A + 97;
                if (Key.isDown(Key.LSHIFT) || Key.isDown(Key.RSHIFT)) c -= 32;
                ch = String.fromCharCode(c);
            } else if (k >= Key.NUMBER_0 && k <= Key.NUMBER_9) {
                ch = String.fromCharCode(k - Key.NUMBER_0 + 48);
            } else if (k == 189 || k == Key.NUMPAD_SUB) {
                ch = Key.isDown(Key.LSHIFT) || Key.isDown(Key.RSHIFT) ? "_" : "-";
            } else if (k == 190 || k == Key.NUMPAD_DOT) {
                ch = ".";
            } else if (k == Key.SPACE) {
                ch = " ";
            }
            if (ch != null) {
                cur += ch; hvFormFields.set(hvFormActive, cur); if (disp != null) disp.text = cur;
                return;
            }
        }
    }

    function handleModalInput() {
        for (k in 0...256) {
            if (!Key.isPressed(k)) continue;
            if (k == Key.ENTER) {
                var cb = modalCallback;
                var v = modalValue;
                closeModal();
                if (cb != null) cb(v);
                return;
            }
            if (k == Key.ESCAPE) { closeModal(); return; }
            if (k == Key.BACKSPACE) {
                if (modalValue.length > 0) { modalValue = modalValue.substr(0, modalValue.length - 1); if (modalDisplay != null) modalDisplay.text = modalValue; }
                return;
            }
            var ch:String = null;
            if (k >= Key.A && k <= Key.Z) {
                var c = k - Key.A + 97;
                if (Key.isDown(Key.LSHIFT) || Key.isDown(Key.RSHIFT)) c -= 32;
                ch = String.fromCharCode(c);
            } else if (k >= Key.NUMBER_0 && k <= Key.NUMBER_9) {
                ch = String.fromCharCode(k - Key.NUMBER_0 + 48);
            } else if (k == 189 || k == Key.NUMPAD_SUB) {
                ch = Key.isDown(Key.LSHIFT) || Key.isDown(Key.RSHIFT) ? "_" : "-";
            } else if (k == 190 || k == Key.NUMPAD_DOT) {
                ch = ".";
            } else if (k == Key.SPACE) {
                ch = " ";
            }
            if (ch != null) {
                modalValue += ch; if (modalDisplay != null) modalDisplay.text = modalValue;
                return;
            }
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
            else if (screen == "hosts") hostScrollY = ns;
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

    // ═══════════════════════════════════════
    //  MIKROTIK
    // ═══════════════════════════════════════
    function showMikroTik() {
        screen = "mikrotik";
        s2d.removeChildren();
        mtEditingHost = false;
        mtEditingUser = false;
        mtEditingPass = false;
        modalActive = false;

        // Carrega credenciais salvas (DPAPI)
        mtLoadCreds();

        drawHeader("MikroTik Manager");
        makeButton("< Voltar", 10, 46, 80, 22, function() { showComputerList(); });

        var lbl = makeText("Host:", 1.3, 0xCDD6F4); lbl.x = 100; lbl.y = 48;
        mtHostBorder = new Graphics(s2d);
        drawFieldBox(mtHostBorder, 145, 46, 200, 22, false);
        mtHostDisplay = makeText(mtHost, 1.2, 0xFFFFFF);
        mtHostDisplay.x = 150; mtHostDisplay.y = 48;
        var hArea = new Interactive(200, 22, s2d);
        hArea.x = 145; hArea.y = 46; hArea.cursor = Button;
        hArea.onClick = function(_) { mtEditingHost = true; drawFieldBox(mtHostBorder, 145, 46, 200, 22, true); };

        makeButton("Testar Portas", 355, 46, 110, 22, function() {
            mtEditingHost = false; drawFieldBox(mtHostBorder, 145, 46, 200, 22, false); mtTestPorts();
        });
        makeButton("Abrir WinBox", 475, 46, 110, 22, function() { mtOpenWinbox(); });
        makeButton("Abrir WebFig", 595, 46, 110, 22, function() { mtOpenWebfig(); });
        makeButton("Identificar", 715, 46, 100, 22, function() { mtIdentify(); });

        var info = makeText("Portas: WinBox=" + mtWinboxPort + "  WebFig=" + mtWebfigPort + "  Telnet=" + mtTelnetPort, 1.0, 0x6C7086);
        info.x = 825; info.y = 50;

        mtStatusText = makeText("Pronto. Clique em Testar Portas para verificar conectividade.", 1.2, 0xF9E2AF);
        mtStatusText.x = 10; mtStatusText.y = 76;

        // Segunda linha: credenciais MK (opcional, se vazio usa login AD)
        var uLbl = makeText("MK User:", 1.2, 0xCDD6F4); uLbl.x = 10; uLbl.y = 102;
        mtUserBorder = new Graphics(s2d);
        drawFieldBox(mtUserBorder, 85, 100, 170, 22, false);
        mtUserDisplay = makeText(mtUser, 1.1, 0xFFFFFF);
        mtUserDisplay.x = 90; mtUserDisplay.y = 103;
        var uArea = new Interactive(170, 22, s2d);
        uArea.x = 85; uArea.y = 100; uArea.cursor = Button;
        uArea.onClick = function(_) {
            mtEditingUser = true; mtEditingPass = false; mtEditingHost = false;
            drawFieldBox(mtUserBorder, 85, 100, 170, 22, true);
            drawFieldBox(mtPassBorder, 320, 100, 170, 22, false);
        };

        var pLbl = makeText("MK Pass:", 1.2, 0xCDD6F4); pLbl.x = 265; pLbl.y = 102;
        mtPassBorder = new Graphics(s2d);
        drawFieldBox(mtPassBorder, 320, 100, 170, 22, false);
        mtPassDisplay = makeText(maskPass(mtPass), 1.1, 0xFFFFFF);
        mtPassDisplay.x = 325; mtPassDisplay.y = 103;
        var pArea = new Interactive(170, 22, s2d);
        pArea.x = 320; pArea.y = 100; pArea.cursor = Button;
        pArea.onClick = function(_) {
            mtEditingPass = true; mtEditingUser = false; mtEditingHost = false;
            drawFieldBox(mtPassBorder, 320, 100, 170, 22, true);
            drawFieldBox(mtUserBorder, 85, 100, 170, 22, false);
        };

        var credHint = makeText("(deixe vazio para usar login AD)", 0.95, 0x6C7086);
        credHint.x = 500; credHint.y = 104;

        makeButton("Salvar cred", 720, 100, 95, 22, function() { mtSaveCreds(); });
        makeButton("Esquecer", 820, 100, 80, 22, function() { mtClearCreds(); });

        mtResultsLayer = new h2d.Object(s2d);
        mtResultsLayer.x = 10; mtResultsLayer.y = 135;
    }

    function mtLoadCreds() {
        var result = runBridge("mt-load-creds", {});
        if (result == null) return;
        if (Reflect.field(result, "status") != "ok") return;
        var d:Dynamic = Reflect.field(result, "data");
        if (Reflect.field(d, "found") == true) {
            var h = Std.string(Reflect.field(d, "host"));
            if (h != null && h.length > 0) mtHost = h;
            mtUser = Std.string(Reflect.field(d, "username"));
            mtPass = Std.string(Reflect.field(d, "password"));
        }
    }

    function mtSaveCreds() {
        if (mtUser.length == 0) { setMtStatus("Preencha MK User antes de salvar", 0xF38BA8); return; }
        var result = runBridge("mt-save-creds", { host: mtHost, username: mtUser, password: mtPass });
        if (result == null) { setMtStatus("Bridge falhou", 0xF38BA8); return; }
        if (Reflect.field(result, "status") == "ok") {
            setMtStatus("Credenciais salvas (DPAPI, criptografadas por usuario Windows)", 0xA6E3A1);
        } else {
            setMtStatus("Erro: " + Std.string(Reflect.field(result, "message")), 0xF38BA8);
        }
    }

    function mtClearCreds() {
        var result = runBridge("mt-clear-creds", {});
        mtUser = ""; mtPass = "";
        if (mtUserDisplay != null) mtUserDisplay.text = "";
        if (mtPassDisplay != null) mtPassDisplay.text = "";
        setMtStatus("Credenciais apagadas", 0xF9E2AF);
    }

    function maskPass(s:String):String {
        var out = "";
        for (i in 0...s.length) out += "*";
        return out;
    }

    function setMtStatus(msg:String, color:Int) {
        if (mtStatusText != null) { mtStatusText.text = msg; mtStatusText.textColor = color; }
    }

    function mtTestPorts() {
        if (mtHost.length == 0) { setMtStatus("Host vazio", 0xF38BA8); return; }
        setMtStatus("Testando " + mtHost + "...", 0xF9E2AF);
        var ports = [22, 23, 80, 443, 2000, 8291, 8728, 8729, mtTelnetPort, mtWebfigPort, mtWinboxPort];
        var result = runBridge("mt-test-ports", { host: mtHost, ports: ports });
        if (result == null) { setMtStatus("Bridge falhou", 0xF38BA8); return; }
        if (Reflect.field(result, "status") != "ok") {
            setMtStatus("Erro: " + Std.string(Reflect.field(result, "message")), 0xF38BA8); return;
        }
        var data:Dynamic = Reflect.field(result, "data");
        var portsD:Array<Dynamic> = cast Reflect.field(data, "ports");
        mtResultsLayer.removeChildren();
        var y:Float = 0;
        var hdr = new Graphics(mtResultsLayer);
        hdr.beginFill(0x45475A); hdr.drawRect(0, y, 400, 22); hdr.endFill();
        addLayerText(mtResultsLayer, "Porta", 10, y + 3, 1.2, 0xCDD6F4);
        addLayerText(mtResultsLayer, "Status", 120, y + 3, 1.2, 0xCDD6F4);
        addLayerText(mtResultsLayer, "Servico", 240, y + 3, 1.2, 0xCDD6F4);
        y += 24;
        var labels = new Map<Int, String>();
        labels.set(22, "SSH"); labels.set(23, "Telnet"); labels.set(80, "WebFig HTTP");
        labels.set(443, "WebFig HTTPS"); labels.set(2000, "Bandwidth Test");
        labels.set(8291, "WinBox"); labels.set(8728, "API"); labels.set(8729, "API-SSL");
        labels.set(mtTelnetPort, "Telnet custom"); labels.set(mtWebfigPort, "WebFig custom"); labels.set(mtWinboxPort, "WinBox custom");
        var openCount = 0;
        for (i in 0...portsD.length) {
            var r:Dynamic = portsD[i];
            var p:Int = Reflect.field(r, "port");
            var op:Bool = Reflect.field(r, "open") == true;
            if (op) openCount++;
            var rowColor:Int = (i % 2 == 0) ? 0x313244 : 0x1E1E2E;
            var row = new Graphics(mtResultsLayer);
            row.beginFill(rowColor); row.drawRect(0, y, 400, 22); row.endFill();
            addLayerText(mtResultsLayer, Std.string(p), 10, y + 3, 1.1, 0xCDD6F4);
            addLayerText(mtResultsLayer, op ? "ABERTA" : "fechada", 120, y + 3, 1.1, op ? 0xA6E3A1 : 0x6C7086);
            var nm = labels.exists(p) ? labels.get(p) : "-";
            addLayerText(mtResultsLayer, nm, 240, y + 3, 1.0, 0xA6ADC8);
            y += 24;
        }
        setMtStatus(openCount + " portas abertas em " + mtHost, openCount > 0 ? 0xA6E3A1 : 0xF38BA8);
    }

    function mtOpenWinbox() {
        if (mtHost.length == 0) { setMtStatus("Host vazio", 0xF38BA8); return; }
        var useUser = mtUser.length > 0 ? mtUser : usernameStr;
        var usePass = mtUser.length > 0 ? mtPass : passwordStr;
        setMtStatus("Abrindo WinBox em " + mtHost + ":" + mtWinboxPort + " como " + useUser + "...", 0xF9E2AF);
        var result = runBridge("mt-open-winbox", {
            host: mtHost, port: mtWinboxPort,
            username: useUser, password: usePass
        });
        if (result == null) { setMtStatus("Bridge falhou", 0xF38BA8); return; }
        if (Reflect.field(result, "status") == "ok") {
            var d:Dynamic = Reflect.field(result, "data");
            setMtStatus("WinBox aberto: " + Std.string(Reflect.field(d, "target")), 0xA6E3A1);
        } else {
            setMtStatus(Std.string(Reflect.field(result, "message")), 0xF38BA8);
        }
    }

    function mtOpenWebfig() {
        if (mtHost.length == 0) { setMtStatus("Host vazio", 0xF38BA8); return; }
        var result = runBridge("mt-open-webfig", { host: mtHost, port: mtWebfigPort });
        if (result == null) { setMtStatus("Bridge falhou", 0xF38BA8); return; }
        if (Reflect.field(result, "status") == "ok") {
            var d:Dynamic = Reflect.field(result, "data");
            setMtStatus("Navegador aberto: " + Std.string(Reflect.field(d, "opened")), 0xA6E3A1);
        } else {
            setMtStatus(Std.string(Reflect.field(result, "message")), 0xF38BA8);
        }
    }

    function mtIdentify() {
        if (mtHost.length == 0) { setMtStatus("Host vazio", 0xF38BA8); return; }
        setMtStatus("Consultando " + mtHost + "...", 0xF9E2AF);
        var result = runBridge("mt-identity", { host: mtHost, port: mtWebfigPort });
        if (result == null) { setMtStatus("Bridge falhou", 0xF38BA8); return; }
        if (Reflect.field(result, "status") == "ok") {
            var d:Dynamic = Reflect.field(result, "data");
            var srv = Std.string(Reflect.field(d, "server"));
            var tit = Std.string(Reflect.field(d, "title"));
            var st = Reflect.field(d, "status");
            setMtStatus("HTTP " + Std.string(st) + " | Server=" + srv + " | Title=" + tit, 0xA6E3A1);
        } else {
            setMtStatus(Std.string(Reflect.field(result, "message")), 0xF38BA8);
        }
    }

    function handleMtHostInput() {
        for (k in 0...256) {
            if (!Key.isPressed(k)) continue;
            if (k == Key.ENTER) { mtEditingHost = false; drawFieldBox(mtHostBorder, 145, 46, 200, 22, false); mtTestPorts(); return; }
            if (k == Key.ESCAPE) { mtEditingHost = false; drawFieldBox(mtHostBorder, 145, 46, 200, 22, false); return; }
            if (k == Key.BACKSPACE) {
                if (mtHost.length > 0) { mtHost = mtHost.substr(0, mtHost.length - 1); mtHostDisplay.text = mtHost; }
                return;
            }
            if (k >= Key.A && k <= Key.Z) {
                var c = k - Key.A + 97;
                if (Key.isDown(Key.LSHIFT) || Key.isDown(Key.RSHIFT)) c -= 32;
                mtHost += String.fromCharCode(c); mtHostDisplay.text = mtHost; return;
            }
            if (k >= Key.NUMBER_0 && k <= Key.NUMBER_9) {
                mtHost += String.fromCharCode(k - Key.NUMBER_0 + 48); mtHostDisplay.text = mtHost; return;
            }
            if (k == 189 || k == Key.NUMPAD_SUB) { mtHost += "-"; mtHostDisplay.text = mtHost; return; }
            if (k == 190 || k == Key.NUMPAD_DOT) { mtHost += "."; mtHostDisplay.text = mtHost; return; }
        }
    }

    function mtTextChar(k:Int):String {
        var shift = Key.isDown(Key.LSHIFT) || Key.isDown(Key.RSHIFT);
        if (k >= Key.A && k <= Key.Z) {
            var c = k - Key.A + 97; if (shift) c -= 32; return String.fromCharCode(c);
        }
        if (k >= Key.NUMBER_0 && k <= Key.NUMBER_9) {
            if (shift) {
                var syms = [")","!","@","#","$","%","^","&","*","("];
                return syms[k - Key.NUMBER_0];
            }
            return String.fromCharCode(k - Key.NUMBER_0 + 48);
        }
        if (k >= Key.NUMPAD_0 && k <= Key.NUMPAD_9) return String.fromCharCode(k - Key.NUMPAD_0 + 48);
        if (k == 189 || k == Key.NUMPAD_SUB) return shift ? "_" : "-";
        if (k == 190 || k == Key.NUMPAD_DOT) return ".";
        if (k == 191) return shift ? "?" : "/";
        if (k == 186) return shift ? ":" : ";";
        if (k == 187 || k == Key.NUMPAD_ADD) return shift ? "+" : "=";
        if (k == 192) return shift ? "~" : "`";
        if (k == 219) return shift ? "{" : "[";
        if (k == 221) return shift ? "}" : "]";
        if (k == 220) return shift ? "|" : "\\";
        if (k == 222) return shift ? "\"" : "'";
        if (k == 188) return shift ? "<" : ",";
        if (k == Key.SPACE) return " ";
        return null;
    }

    function handleMtUserInput() {
        for (k in 0...256) {
            if (!Key.isPressed(k)) continue;
            if (k == Key.ENTER || k == Key.TAB) { mtEditingUser = false; drawFieldBox(mtUserBorder, 85, 100, 170, 22, false); return; }
            if (k == Key.ESCAPE) { mtEditingUser = false; drawFieldBox(mtUserBorder, 85, 100, 170, 22, false); return; }
            if (k == Key.BACKSPACE) {
                if (mtUser.length > 0) { mtUser = mtUser.substr(0, mtUser.length - 1); mtUserDisplay.text = mtUser; }
                return;
            }
            var ch = mtTextChar(k);
            if (ch != null) { mtUser += ch; mtUserDisplay.text = mtUser; return; }
        }
    }

    function handleMtPassInput() {
        for (k in 0...256) {
            if (!Key.isPressed(k)) continue;
            if (k == Key.ENTER || k == Key.TAB) { mtEditingPass = false; drawFieldBox(mtPassBorder, 320, 100, 170, 22, false); return; }
            if (k == Key.ESCAPE) { mtEditingPass = false; drawFieldBox(mtPassBorder, 320, 100, 170, 22, false); return; }
            if (k == Key.BACKSPACE) {
                if (mtPass.length > 0) { mtPass = mtPass.substr(0, mtPass.length - 1); mtPassDisplay.text = maskPass(mtPass); }
                return;
            }
            var ch = mtTextChar(k);
            if (ch != null) { mtPass += ch; mtPassDisplay.text = maskPass(mtPass); return; }
        }
    }

    static function main() { new Main(); }
}
