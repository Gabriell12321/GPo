import h2d.Text;
import h2d.Graphics;
import h2d.Interactive;
import hxd.Key;
import hxd.res.DefaultFont;

class Main extends hxd.App {
    // Config - auto-detectado do sistema
    static var DOMAIN:String = "";
    static var DEFAULT_USER:String = "";

    // State
    var screen:String = "login";
    var font:h2d.Font;
    var centerX:Float;
    var centerY:Float;

    // Login fields
    var usernameStr:String = "";
    var passwordStr:String = "";
    var activeField:String = "password"; // "username" ou "password"
    var usernameDisplay:Text;
    var passwordDisplay:Text;
    var messageText:Text;
    var usernameBorder:Graphics;
    var passwordBorder:Graphics;

    // Login field positions (stored for click handlers)
    var userFieldY:Float = 0;
    var passFieldY:Float = 0;
    var fieldX:Float = 0;
    var fieldW:Float = 260;
    var fieldH:Float = 26;

    // Dashboard
    var gpoData:Array<Dynamic> = [];
    var gpoTexts:Array<Text> = [];
    var dashLayer:h2d.Object;
    var statusText:Text;
    var scrollY:Float = 0;

    override function init() {
        engine.backgroundColor = 0xFF1E1E2E;
        font = DefaultFont.get();
        centerX = s2d.width / 2;
        centerY = s2d.height / 2;

        // Auto-detectar dominio e usuario do sistema
        try {
            var domResult = new sys.io.Process("powershell", ["-NoProfile", "-Command", "(Get-WmiObject Win32_ComputerSystem).Domain"]);
            var domOut = StringTools.trim(domResult.stdout.readAll().toString());
            domResult.close();
            if (domOut.length > 0 && domOut != "WORKGROUP") DOMAIN = domOut;
        } catch(_:Dynamic) {}
        try {
            var userResult = new sys.io.Process("powershell", ["-NoProfile", "-Command", "$env:USERNAME"]);
            var userOut = StringTools.trim(userResult.stdout.readAll().toString());
            userResult.close();
            if (userOut.length > 0) { DEFAULT_USER = userOut; usernameStr = userOut; }
        } catch(_:Dynamic) {}

        showLogin();
    }

    // ═══════════════════════════════════════
    //  TELA DE LOGIN
    // ═══════════════════════════════════════
    function showLogin() {
        screen = "login";
        s2d.removeChildren();
        activeField = "password";

        fieldX = centerX - 130;

        // Titulo
        var title = makeText("GPO - Sistema de TI", 2, 0x89B4FA);
        title.x = centerX - (title.textWidth * 2) / 2;
        title.y = 50;

        var sub = makeText(DOMAIN.length > 0 ? "Dominio: " + DOMAIN : "Dominio: (nao detectado)", 1.2, 0x6C7086);
        sub.x = centerX - (sub.textWidth * 1.2) / 2;
        sub.y = 90;

        // Label usuario
        var userLabel = makeText("Usuario:", 1.5, 0xCDD6F4);
        userLabel.x = fieldX;
        userLabel.y = 130;

        // Caixa usuario
        userFieldY = 152;
        usernameBorder = new Graphics(s2d);
        drawFieldBox(usernameBorder, fieldX, userFieldY, fieldW, fieldH, false);

        usernameDisplay = makeText(usernameStr, 1.5, 0xFFFFFF);
        usernameDisplay.x = fieldX + 5;
        usernameDisplay.y = userFieldY + 4;

        // Interactive usuario
        var userArea = new Interactive(fieldW, fieldH, s2d);
        userArea.x = fieldX;
        userArea.y = userFieldY;
        userArea.cursor = Button;
        userArea.onClick = function(_) { selectField("username"); };

        // Label senha
        var passLabel = makeText("Senha:", 1.5, 0xCDD6F4);
        passLabel.x = fieldX;
        passLabel.y = 192;

        // Caixa senha
        passFieldY = 214;
        passwordBorder = new Graphics(s2d);
        drawFieldBox(passwordBorder, fieldX, passFieldY, fieldW, fieldH, true);

        passwordDisplay = makeText("", 1.5, 0xFFFFFF);
        passwordDisplay.x = fieldX + 5;
        passwordDisplay.y = passFieldY + 4;

        // Interactive senha
        var passArea = new Interactive(fieldW, fieldH, s2d);
        passArea.x = fieldX;
        passArea.y = passFieldY;
        passArea.cursor = Button;
        passArea.onClick = function(_) { selectField("password"); };

        // Botao Entrar
        makeButton("Entrar", centerX - 50, 258, 100, 28, function() { doLogin(); });

        // Mensagem
        messageText = makeText("", 1.3, 0xF38BA8);
        messageText.y = 300;
    }

    function selectField(field:String) {
        activeField = field;
        drawFieldBox(usernameBorder, fieldX, userFieldY, fieldW, fieldH, field == "username");
        drawFieldBox(passwordBorder, fieldX, passFieldY, fieldW, fieldH, field == "password");
    }

    function doLogin() {
        messageText.textColor = 0xF9E2AF;
        messageText.text = "Autenticando...";
        messageText.x = centerX - (messageText.textWidth * 1.3) / 2;

        var result = runBridge("auth", {
            domain: DOMAIN,
            username: usernameStr,
            password: passwordStr
        });

        if (result == null) {
            messageText.textColor = 0xF38BA8;
            messageText.text = "Erro: falha ao executar bridge";
            messageText.x = centerX - (messageText.textWidth * 1.3) / 2;
            return;
        }

        var status:String = Reflect.field(result, "status");
        if (status == "ok") {
            var data:Dynamic = Reflect.field(result, "data");
            var auth:Bool = Reflect.field(data, "authenticated");
            if (auth) {
                showDashboard();
            } else {
                messageText.textColor = 0xF38BA8;
                messageText.text = "Credenciais invalidas!";
                messageText.x = centerX - (messageText.textWidth * 1.3) / 2;
                passwordStr = "";
                updateFieldDisplay();
            }
        } else {
            var msg:String = Reflect.field(result, "message");
            messageText.textColor = 0xF38BA8;
            messageText.text = msg != null ? msg : "Erro desconhecido";
            messageText.x = centerX - (messageText.textWidth * 1.3) / 2;
        }
    }

    // ═══════════════════════════════════════
    //  DASHBOARD
    // ═══════════════════════════════════════
    function showDashboard() {
        screen = "dashboard";
        s2d.removeChildren();
        scrollY = 0;

        dashLayer = new h2d.Object(s2d);

        // Header
        var header = new Graphics(s2d);
        header.beginFill(0x313244);
        header.drawRect(0, 0, s2d.width, 40);
        header.endFill();

        var htitle = makeText("GPO Manager - " + DOMAIN, 1.5, 0x89B4FA);
        htitle.x = 10;
        htitle.y = 10;

        var logoutBtn = makeText("[Sair]", 1.3, 0xF38BA8);
        logoutBtn.x = s2d.width - 60;
        logoutBtn.y = 12;
        var logoutArea = new Interactive(50, 20, s2d);
        logoutArea.x = s2d.width - 60;
        logoutArea.y = 12;
        logoutArea.cursor = Button;
        logoutArea.onClick = function(_) {
            passwordStr = "";
            showLogin();
        };

        // Status
        statusText = makeText("Carregando GPOs...", 1.2, 0xF9E2AF);
        statusText.x = 10;
        statusText.y = 50;

        // Carregar GPOs
        loadGPOs();
    }

    function loadGPOs() {
        var result = runBridge("list-gpos", {domain: DOMAIN});

        if (result == null) {
            statusText.text = "Erro: falha ao executar bridge PowerShell";
            statusText.textColor = 0xF38BA8;
            return;
        }

        var status:String = Reflect.field(result, "status");
        if (status == "ok") {
            var data:Dynamic = Reflect.field(result, "data");
            if (data != null && Std.isOfType(data, Array)) {
                gpoData = cast data;
                statusText.text = "GPOs encontradas: " + gpoData.length;
                statusText.textColor = 0xA6E3A1;
                renderGPOList();
            } else {
                statusText.text = "Nenhuma GPO encontrada";
                statusText.textColor = 0xF9E2AF;
            }
        } else {
            var msg:String = Reflect.field(result, "message");
            statusText.text = msg != null ? msg : "Erro ao listar GPOs";
            statusText.textColor = 0xF38BA8;
        }
    }

    function renderGPOList() {
        // Limpar lista anterior
        if (dashLayer != null) dashLayer.removeChildren();

        var y:Float = 0;
        var w = s2d.width - 20;

        // Cabecalho da tabela
        var hdr = new Graphics(dashLayer);
        hdr.beginFill(0x45475A);
        hdr.drawRect(0, y, w, 22);
        hdr.endFill();

        addTableText("Nome", 10, y + 2, 1.2, 0xCDD6F4);
        addTableText("Status", 280, y + 2, 1.2, 0xCDD6F4);
        addTableText("Criada", 420, y + 2, 1.2, 0xCDD6F4);
        addTableText("Modificada", 560, y + 2, 1.2, 0xCDD6F4);
        y += 24;

        for (i in 0...gpoData.length) {
            var gpo:Dynamic = gpoData[i];
            var color:Int = (i % 2 == 0) ? 0x313244 : 0x1E1E2E;

            var row = new Graphics(dashLayer);
            row.beginFill(color);
            row.drawRect(0, y, w, 20);
            row.endFill();

            var name:String = Reflect.field(gpo, "name");
            var gpoStatus:String = Reflect.field(gpo, "status");
            var created:String = Reflect.field(gpo, "created");
            var modified:String = Reflect.field(gpo, "modified");

            addTableText(name != null ? name : "?", 10, y + 2, 1.0, 0xCDD6F4);

            var statusColor:Int = 0xA6E3A1;
            if (gpoStatus != null && gpoStatus != "AllSettingsEnabled") statusColor = 0xF9E2AF;
            addTableText(gpoStatus != null ? gpoStatus : "?", 280, y + 2, 1.0, statusColor);

            addTableText(created != null ? created : "?", 420, y + 2, 1.0, 0x6C7086);
            addTableText(modified != null ? modified : "?", 560, y + 2, 1.0, 0x6C7086);

            y += 22;
        }

        dashLayer.x = 10;
        dashLayer.y = 70;
    }

    function addTableText(str:String, x:Float, y:Float, scale:Float, color:Int) {
        var t = new Text(font, dashLayer);
        t.text = str;
        t.setScale(scale);
        t.textColor = color;
        t.x = x;
        t.y = y;
    }

    // ═══════════════════════════════════════
    //  BRIDGE POWERSHELL
    // ═══════════════════════════════════════
    function runBridge(cmd:String, args:Dynamic):Dynamic {
        try {
            var inputFile = "scripts/_cmd.json";
            var outputFile = "scripts/_result.json";

            // Escrever comando
            var cmdObj:Dynamic = {cmd: cmd, args: args};
            var jsonStr = haxe.Json.stringify(cmdObj);
            sys.io.File.saveContent(inputFile, jsonStr);

            // Executar bridge
            var exitCode = Sys.command("powershell", [
                "-NoProfile", "-ExecutionPolicy", "Bypass",
                "-File", "scripts/bridge.ps1",
                "-InputFile", inputFile,
                "-OutputFile", outputFile
            ]);

            // Limpar arquivo de comando (contem senha em auth)
            try { sys.FileSystem.deleteFile(inputFile); } catch(_:Dynamic) {}

            if (exitCode != 0 && exitCode != 1) {
                try { sys.FileSystem.deleteFile(outputFile); } catch(_:Dynamic) {}
                return null;
            }

            // Ler resultado
            if (sys.FileSystem.exists(outputFile)) {
                var content = sys.io.File.getContent(outputFile);
                try { sys.FileSystem.deleteFile(outputFile); } catch(_:Dynamic) {}
                return haxe.Json.parse(content);
            }
            return null;
        } catch(e:Dynamic) {
            return null;
        }
    }

    // ═══════════════════════════════════════
    //  INPUT DE TECLADO
    // ═══════════════════════════════════════
    override function update(dt:Float) {
        if (screen != "login") {
            // Scroll no dashboard
            if (Key.isPressed(Key.MOUSE_WHEEL_DOWN) && dashLayer != null) {
                scrollY -= 20;
                dashLayer.y = 70 + scrollY;
            }
            if (Key.isPressed(Key.MOUSE_WHEEL_UP) && dashLayer != null) {
                scrollY += 20;
                if (scrollY > 0) scrollY = 0;
                dashLayer.y = 70 + scrollY;
            }
            return;
        }

        for (k in 0...256) {
            if (Key.isPressed(k)) {
                if (k == Key.TAB) {
                    selectField((activeField == "username") ? "password" : "username");
                } else if (k == Key.BACKSPACE) {
                    if (activeField == "username" && usernameStr.length > 0) {
                        usernameStr = usernameStr.substr(0, usernameStr.length - 1);
                    } else if (activeField == "password" && passwordStr.length > 0) {
                        passwordStr = passwordStr.substr(0, passwordStr.length - 1);
                    }
                    updateFieldDisplay();
                } else if (k == Key.ENTER) {
                    doLogin();
                } else if (k >= Key.A && k <= Key.Z) {
                    var c = k - Key.A + 97;
                    if (Key.isDown(Key.LSHIFT) || Key.isDown(Key.RSHIFT)) c -= 32;
                    appendChar(String.fromCharCode(c));
                } else if (k >= Key.NUMBER_0 && k <= Key.NUMBER_9) {
                    var ch = String.fromCharCode(k - Key.NUMBER_0 + 48);
                    if (Key.isDown(Key.LSHIFT) || Key.isDown(Key.RSHIFT)) {
                        // Caracteres especiais com shift
                        var specials = [")", "!", "@", "#", "$", "%", "^", "&", "*", "("];
                        ch = specials[k - Key.NUMBER_0];
                    }
                    appendChar(ch);
                } else if (k == Key.NUMPAD_DOT || k == 190) {
                    appendChar(".");
                } else if (k == 189 || k == Key.NUMPAD_SUB) {
                    if (Key.isDown(Key.LSHIFT) || Key.isDown(Key.RSHIFT))
                        appendChar("_");
                    else
                        appendChar("-");
                }
            }
        }
    }

    function appendChar(c:String) {
        if (activeField == "username")
            usernameStr += c;
        else
            passwordStr += c;
        updateFieldDisplay();
    }

    function updateFieldDisplay() {
        if (usernameDisplay != null)
            usernameDisplay.text = usernameStr;
        if (passwordDisplay != null) {
            var masked = "";
            for (_ in 0...passwordStr.length) masked += "*";
            passwordDisplay.text = masked;
        }
    }

    // ═══════════════════════════════════════
    //  HELPERS DE UI
    // ═══════════════════════════════════════
    function makeText(str:String, scale:Float, color:Int):Text {
        var t = new Text(font, s2d);
        t.text = str;
        t.setScale(scale);
        t.textColor = color;
        return t;
    }

    function makeButton(label:String, x:Float, y:Float, w:Float, h:Float, onClick:Void->Void) {
        var bg = new Graphics(s2d);
        bg.beginFill(0x89B4FA);
        bg.drawRect(x, y, w, h);
        bg.endFill();

        var t = makeText(label, 1.5, 0x1E1E2E);
        t.x = x + w / 2 - (t.textWidth * 1.5) / 2;
        t.y = y + 4;

        var area = new Interactive(w, h, s2d);
        area.x = x;
        area.y = y;
        area.cursor = Button;
        area.onClick = function(_) { onClick(); };
        area.onOver = function(_) {
            bg.clear();
            bg.beginFill(0xB4D0FA);
            bg.drawRect(x, y, w, h);
            bg.endFill();
        };
        area.onOut = function(_) {
            bg.clear();
            bg.beginFill(0x89B4FA);
            bg.drawRect(x, y, w, h);
            bg.endFill();
        };
    }

    function drawFieldBox(g:Graphics, x:Float, y:Float, w:Float, h:Float, active:Bool) {
        g.clear();
        g.beginFill(0x313244);
        g.drawRect(x, y, w, h);
        g.endFill();
        g.lineStyle(1, active ? 0x89B4FA : 0x45475A);
        g.drawRect(x, y, w, h);
    }

    static function main() {
        new Main();
    }
}
