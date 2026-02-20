Using Namespace System
Using Namespace System.Drawing
Using Namespace System.Windows
Using Namespace System.Windows.Forms
Using Namespace System.Management.Automation

# ============================================================================
# Create365User.ps1 - MSP Edition
# ============================================================================
# Maakt een nieuwe gebruiker aan in On-Premises AD of Cloud Only (Office 365)
# Geen INI bestand nodig - alles wordt ingevuld/opgehaald tijdens het draaien.
#
# Flow:
#   1. Kies omgeving: On-Premises AD of Cloud Only (365)
#   2. On-Prem: Vul AD credentials in -> MS Graph login (interactief)
#   3. Cloud Only: MS Graph login (interactief)
#   4. Licenties, domeinen, tenant info worden automatisch opgehaald
#   5. Vul gebruikersgegevens in via het formulier
#   6. Gebruiker wordt aangemaakt
# ============================================================================

$DebugPreference = 'Continue'

$PSPolicy = 'RemoteSigned'
$ExecPolicy = (Get-ExecutionPolicy -Scope CurrentUser)
if ($ExecPolicy -ne $PSPolicy) {
    try {
        Set-ExecutionPolicy -ExecutionPolicy "$PSPolicy" -Scope CurrentUser
    } catch {
        Write-Debug "Dit script vereist ExecutionPolicy $PSPolicy"
        exit 1
    }
}

function Get-CurrentPath {
    $currentPath = $PSScriptRoot
    if (!$currentPath) { $currentPath = Split-Path $pseditor.GetEditorContext().CurrentFile.Path -ErrorAction SilentlyContinue }
    if (!$currentPath) { $currentPath = Split-Path $psISE.CurrentFile.FullPath -ErrorAction SilentlyContinue }
    return $currentPath + '\'
}

$cfolder = Get-CurrentPath
$ScriptName = "$($MyInvocation.MyCommand.Name)"
$LogFileName = $ScriptName.Replace(".ps1", ".log")
$LOGFile = "$($cFolder)$($LogFileName)"
if ([System.IO.File]::Exists($LOGFile)) {
    try { Remove-Item -Path $LOGFile -Force -ErrorAction SilentlyContinue } catch {}
}

# ============================================================================
# .NET Assemblies laden
# ============================================================================
try {
    Add-Type -AssemblyName System.Windows.Forms -WarningAction SilentlyContinue
    Add-Type -AssemblyName System.Drawing -WarningAction SilentlyContinue
    Add-Type -AssemblyName PresentationCore, PresentationFramework -WarningAction SilentlyContinue
} catch {
    Write-Error "Kan .NET assemblies niet laden."
    exit
}

[System.Windows.Forms.Application]::EnableVisualStyles()

# ============================================================================
# C# Helpers: Shadow Form + Round Buttons
# ============================================================================
$Shadow = @'
using System;
using System.Windows;
using System.Windows.Forms;
namespace Program
{
    public partial class Shadow: Form
    {
        protected override CreateParams CreateParams
        {
            get
            {
                const int CS_DROPSHADOW = 0x20000;
                CreateParams cp = base.CreateParams;
                cp.ClassStyle |= CS_DROPSHADOW;
                return cp;
            }
        }
    }
}
'@
try { Add-Type -TypeDefinition $Shadow -Language CSharp -WarningAction SilentlyContinue -ReferencedAssemblies System, System.Windows, System.Windows.Forms, System.ComponentModel.Primitives } catch {}

$code = @'
[System.Runtime.InteropServices.DllImport("gdi32.dll")]
public static extern IntPtr CreateRoundRectRgn(int nLeftRect, int nTopRect, int nRightRect, int nBottomRect, int nWidthEllipse, int nHeightEllipse);
'@
try { $Win32Helpers = Add-Type -MemberDefinition $code -Name "Win32Helpers" -WarningAction SilentlyContinue -PassThru } catch {}

# ============================================================================
# GUI Variabelen
# ============================================================================
$Font14B = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$Font12 = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Regular)
$Font10B = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$Font10 = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Regular)
$Font9 = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$Font8 = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Regular)
$Font7 = New-Object System.Drawing.Font("Segoe UI", 7, [System.Drawing.FontStyle]::Regular)
$Tick_Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)

$BackColor = [System.Drawing.Color]::FromArgb(42, 43, 47)
$ForeColor = [System.Drawing.Color]::FromKnownColor("ControlLight")

$ButtonFont = $Font10B
$ButtonBorderColor = [System.Drawing.Color]::FromArgb(163, 42, 64)
$ButtonColor = [System.Drawing.Color]::FromArgb(163, 42, 64)
$ButtonMouseDownColor = [System.Drawing.Color]::FromArgb(237, 56, 89)
$ButtonMouseOverColor = [System.Drawing.Color]::FromArgb(197, 35, 65)

$TextBoxBackColor = [System.Drawing.Color]::Black
$TextBoxFont = $Font9
$Form_Font = $Font10

# ============================================================================
# Globale variabelen
# ============================================================================
$global:ADusers = $null
$global:UPNS = $null
$global:EnvironmentMode = $null        # "OnPrem" of "CloudOnly"
$global:ADCredential = $null
$global:ADServer = $null
$global:GraphConnected = $false
$global:TenantDomains = @()
$global:TenantLicenses = @()
$global:PrimaryDomain = ""
$global:CompanyName = ""

# ============================================================================
# HTML voor email
# ============================================================================
$Head = @'
<div>
<style>
  body { font-family: Arial; font-size: 8pt; color: #4C607B; }
  table { cellpadding: 0; callspacing:0; margin: 0px; padding: 5px; border-spacing: 0px; border: 1px solid #D3D3D3; border-collapse: collapse; font-size: 1.2em; text-align: left; }
  th { border: 1px solid #D3D3D3; padding: 5px; background-color: #003366; color: #ffffff; }
  td { border: 1px solid #D3D3D3; padding: 5px; color: #000000; }
  tr:nth-child(even) { background-color: #f2f2f2; }
</style>
</div>
'@

# ============================================================================
# KeyPress filter
# ============================================================================
$invalidKeys = @('next','multiply','OemBackslash','OemClear','OemCloseBrackets','Oemcomma','OemOpenBrackets','OemPipe','Oemplus','OemQuestion','OemQuotes','OemSemicolon','Oemtilde','Oem5','Oem6','d3','d4','d5','d6','d7','d8','d9','divide')

# ============================================================================
# Wachtwoord woordenlijsten
# ============================================================================
$nouns = @('able','account','achieve','achiever','acoustics','act','action','activity','actor','addition','adjustment','advice','afternoon','agreement','air','airplane','airport','alarm','amount','anger','angle','animal','answer','ant','apparatus','apple','appliance','approval','arch','argument','arm','army','art','attack','attempt','attention','attraction','aunt','authority','baby','back','badge','bag','balance','ball','balloon','banana','band','base','baseball','basin','basket','bat','bath','battle','bead','beam','bean','bear','beast','bed','bedroom','bee','beef','beetle','beginner','belief','bell','berry','bike','bird','birth','birthday','bit','bite','blade','blood','blow','board','boat','body','bomb','bone','book','boot','border','bottle','box','boy','brain','brake','branch','brass','bread','breakfast','breath','brick','bridge','brother','brush','bubble','bucket','building','bulb','burn','burst','business','butter','button','cabbage','cable','cake','calculator','calendar','camera','camp','can','cannon','canvas','cap','car','card','care','carpenter','carriage','cart','cast','cat','cattle','cause','cave','cellar','cent','chain','chair','chalk','chance','change','channel','cheese','cherry','chess','chicken','children','chin','church','circle','class','clock','cloth','cloud','club','coach','coal','coast','coat','coil','collar','color','comb','comfort','committee','company','competition','condition','connection','control','cook','copper','copy','cord','cork','corn','country','cover','cow','crack','crayon','cream','creature','credit','crime','crowd','crown','crush','cry','cup','current','curtain','curve','cushion','dad','daughter','day','death','debt','decision','deer','degree','design','desire','desk','detail','development','dinner','direction','dirt','discovery','disease','distance','division','dock','doctor','dog','doll','donkey','door','drain','drawer','dress','drink','driving','drop','drug','drum','duck','dust','ear','earth','edge','education','effect','egg','elbow','end','engine','error','event','example','exchange','existence','experience','expert','eye','face','fact','fall','family','fan','farm','farmer','father','fear','feast','feather','feeling','feet','fiction','field','fight','finger','fire','fish','flag','flame','flavor','flesh','flight','flock','floor','flower','fly','fog','fold','food','foot','force','fork','form','frame','friction','friend','frog','front','fruit','fuel','furniture','game','garden','gate','ghost','giraffe','girl','glass','glove','glue','goat','gold','goose','government','grade','grain','grape','grass','grip','ground','group','growth','guide','guitar','gun','hair','hall','hammer','hand','harbor','harmony','hat','head','health','heart','heat','help','hen','hill','history','hole','holiday','home','honey','hook','hope','horn','horse','hose','hospital','hour','house','humor','ice','idea','impulse','income','increase','industry','ink','insect','instrument','insurance','interest','invention','iron','island','jail','jam','jar','jelly','jewel','join','joke','journey','judge','juice','jump','kettle','key','kick','kiss','kite','kitten','knee','knife','knot','knowledge','lake','lamp','land','language','laugh','lawyer','lead','leaf','learning','leather','leg','letter','lettuce','level','library','light','limit','line','linen','lip','liquid','list','lock','look','loss','love','lumber','lunch','machine','magic','maid','mailbox','man','manager','map','marble','mark','market','mask','mass','match','meal','measure','meat','meeting','memory','metal','middle','milk','mind','mine','minister','mint','minute','mist','mitten','money','monkey','month','moon','morning','mother','motion','mountain','mouth','move','muscle','music','nail','name','nation','neck','need','needle','nerve','nest','net','news','night','noise','north','nose','note','notebook','number','nut','ocean','offer','office','oil','operation','opinion','orange','order','organization','oven','owl','owner','page','pail','pain','paint','pan','pancake','paper','parcel','parent','park','part','partner','party','passenger','paste','payment','peace','pear','pen','pencil','person','pest','pet','pickle','picture','pie','pig','pin','pipe','place','plane','plant','plastic','plate','play','playground','pleasure','plot','pocket','point','poison','police','pollution','popcorn','position','pot','potato','powder','power','price','print','prison','process','produce','profit','property','protest','pull','pump','punishment','purpose','push','quarter','queen','question','rabbit','rail','railway','rain','rake','range','rat','rate','ray','reaction','reading','reason','record','regret','relation','religion','request','respect','rest','reward','rhythm','rice','riddle','rifle','ring','river','road','rock','rod','roll','roof','room','root','rose','route','rule','run','sack','sail','salt','sand','scale','scarf','scene','scent','school','science','scissors','screw','sea','seat','seed','selection','sense','shade','shake','shame','shape','sheep','sheet','shelf','ship','shirt','shock','shoe','shop','show','side','sign','silk','silver','sink','sister','size','skate','skin','skirt','sky','sleep','slip','slope','smell','smile','smoke','snail','snake','sneeze','snow','soap','society','sock','soda','sofa','son','song','sort','sound','soup','space','spade','spark','sponge','spoon','spot','spring','spy','square','squirrel','stage','stamp','star','start','statement','station','steam','steel','stem','step','stew','stick','stitch','stocking','stomach','stone','stop','store','story','stove','stranger','straw','stream','street','stretch','string','structure','substance','sugar','suggestion','suit','summer','sun','support','surprise','sweater','swim','swing','system','table','tail','talk','tank','taste','tax','teaching','team','teeth','temper','tent','territory','test','texture','theory','thing','thought','thread','throat','throne','thumb','thunder','ticket','tiger','time','tin','title','toad','toe','tomatoes','tongue','tooth','toothbrush','top','touch','town','toy','trade','trail','train','transport','tray','treatment','tree','trick','trip','trouble','truck','tub','turkey','turn','twig','twist','umbrella','uncle','unit','use','vacation','value','van','vase','vegetable','vein','verse','vessel','vest','view','visitor','voice','volcano','voyage','walk','wall','war','wash','waste','watch','water','wave','wax','way','wealth','weather','week','weight','wheel','whip','whistle','wilderness','wind','window','wine','wing','winter','wire','wish','woman','wood','wool','word','work','worm','wound','wren','wrench','wrist','writer','writing','yard','yarn','year','zebra','zinc','zipper','zoo')
$verbs = @('abide','accelerate','accept','accomplish','achieve','acquire','adapt','add','address','admire','admit','adopt','advise','afford','agree','alert','allow','amuse','analyze','announce','answer','apologize','appear','applaud','appoint','appreciate','approve','argue','arrange','arrive','ask','assemble','assess','assist','attach','attack','attempt','attend','attract','avoid','awake','back','bake','balance','ban','bang','bat','bathe','battle','beam','bear','beat','become','beg','begin','behave','belong','bend','bet','bid','bind','bite','bleach','bleed','bless','blind','blink','blow','blush','boast','boil','bolt','bomb','book','bore','borrow','bounce','bow','box','brake','branch','break','breathe','breed','bring','broadcast','brush','bubble','budget','build','bump','burn','burst','bury','bust','buy','calculate','call','camp','care','carry','carve','cast','catch','cause','challenge','change','charge','chart','chase','check','cheer','chew','choke','choose','chop','claim','clap','clarify','clean','clear','cling','clip','close','coach','coil','collect','color','comb','come','command','communicate','compare','compete','compile','complain','complete','compose','compute','concentrate','concern','conclude','conduct','confess','confront','confuse','connect','conserve','consider','consist','construct','consult','contain','continue','contract','control','convert','coordinate','copy','correct','cost','cough','count','cover','crack','crash','crawl','create','cross','crush','cry','cure','curl','curve','cut','cycle','damage','dance','dare','deal','decay','decide','decorate','define','delay','delegate','delight','deliver','demonstrate','depend','describe','deserve','design','destroy','detect','determine','develop','diagnose','dig','direct','disagree','discover','display','distribute','dive','divide','double','doubt','draft','drag','drain','draw','dream','dress','drink','drip','drive','drop','drum','dry','dust','earn','eat','educate','eliminate','employ','encourage','end','endure','enforce','engineer','enhance','enjoy','ensure','enter','entertain','escape','establish','estimate','evaluate','examine','exceed','excite','excuse','exercise','exhibit','exist','expand','expect','experiment','explain','explode','express','extend','extract','face','fade','fail','fasten','fear','feed','feel','fence','fetch','fight','file','fill','film','finalize','finance','find','fire','fit','fix','flash','flee','float','flood','flow','fly','fold','follow','fool','forbid','force','forecast','forget','forgive','form','frame','freeze','frighten','gather','gaze','generate','get','give','glow','glue','go','govern','grab','graduate','greet','grin','grind','grip','grow','guarantee','guard','guess','guide','hammer','hand','handle','hang','happen','harm','hate','head','heal','hear','heat','help','hide','hit','hold','hook','hop','hope','hover','hug','hunt','hurry','hurt','identify','ignore','illustrate','imagine','implement','impress','improve','include','increase','influence','inform','inject','innovate','inspect','inspire','install','instruct','insure','integrate','intend','interest','interpret','interrupt','interview','introduce','invent','investigate','invite','irritate','jail','jam','jog','join','joke','judge','jump','justify','keep','kick','kill','kiss','kneel','knit','knock','know','label','land','last','laugh','launch','lay','lead','lean','leap','learn','leave','lend','let','level','license','lick','lie','lift','light','lighten','like','list','listen','live','load','locate','lock','long','look','lose','love','maintain','make','manage','manufacture','map','march','mark','market','marry','match','matter','mean','measure','meet','melt','memorize','mend','milk','mine','miss','mix','model','modify','monitor','motivate','mourn','move','multiply','name','navigate','need','negotiate','nest','nod','note','notice','number','obey','observe','obtain','occur','offend','offer','open','operate','order','organize','overcome','owe','own','pack','paint','park','part','participate','pass','paste','pat','pause','pay','peel','perceive','perform','permit','persuade','phone','pick','pilot','pinch','place','plan','plant','play','please','plug','point','polish','pop','possess','post','pour','practice','pray','predict','prefer','prepare','present','preserve','press','pretend','prevent','print','process','produce','program','progress','project','promise','promote','propose','protect','prove','provide','pull','pump','punch','punish','purchase','push','put','qualify','question','quit','race','rain','raise','rank','rate','reach','read','realize','reason','receive','recognize','recommend','record','recruit','reduce','refer','reflect','refuse','regret','regulate','reinforce','reject','relate','relax','release','rely','remain','remember','remind','remove','repair','repeat','replace','reply','report','represent','request','rescue','research','resolve','respond','restore','retire','retrieve','return','review','revise','ride','ring','rise','risk','rob','rock','roll','rub','ruin','rule','run','rush','sail','satisfy','save','say','scare','scatter','schedule','scold','scrape','scratch','scream','scrub','seal','search','secure','see','seek','select','sell','send','sense','separate','serve','set','settle','sew','shade','shake','shape','share','shave','shed','shelter','shine','shiver','shock','shoot','shop','show','shrink','shut','sigh','sign','signal','simplify','sing','sink','sip','sit','sketch','ski','skip','slap','sleep','slide','slip','slow','smash','smell','smile','smoke','snatch','sneak','sneeze','snow','soak','solve','sort','sound','spare','spark','speak','specify','speed','spell','spend','spill','spin','split','spoil','spot','spray','spread','spring','squash','squeeze','stain','stamp','stand','stare','start','stay','steal','steer','step','stick','stimulate','sting','stir','stop','store','strap','strengthen','stretch','strike','string','strip','strive','stroke','structure','study','stuff','subtract','succeed','suffer','suggest','summarize','supervise','supply','support','suppose','surprise','surround','suspect','suspend','swear','sweep','swell','swim','swing','switch','symbolize','take','talk','tame','tap','target','taste','teach','tear','tease','tell','tempt','terrify','test','thank','think','throw','tick','tickle','tie','time','tip','tire','touch','tour','tow','trace','trade','train','transfer','transform','translate','transport','trap','travel','treat','tremble','trick','trip','trouble','trust','try','tug','tumble','turn','tutor','twist','type','understand','undertake','unite','unlock','unpack','update','upgrade','upset','use','utilize','vanish','verify','visit','wait','wake','walk','wander','want','warm','warn','wash','waste','watch','water','wave','wear','weave','weep','weigh','welcome','wet','whip','whisper','whistle','win','wind','wink','wipe','wish','withdraw','wonder','work','worry','wrap','wreck','wrestle','wring','write','yawn','yell','zip','zoom')

# ============================================================================
# Helper Functions
# ============================================================================
Function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string] $message,
        [Parameter(Mandatory = $false)][ValidateSet("INFO","WARN","ERROR","SUCCESS")][string] $level = "INFO"
    )
    $timestamp = (Get-Date).toString("dd/MM/yyyy HH:mm:ss")
    $logLine = "$timestamp [$level] - $message"
    Add-Content -Path $LOGFile -Value $logLine
    switch ($level) {
        "ERROR"   { Write-Host $logLine -ForegroundColor Red }
        "WARN"    { Write-Host $logLine -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $logLine -ForegroundColor Green }
        default   { Write-Host $logLine -ForegroundColor Gray }
    }
}

Function Remove_All_Controls($aForm) {
    $Indexes = @()
    Foreach ($control in $aForm.Controls) { $Indexes += $control.TabIndex }
    Foreach ($index in $indexes | Sort-Object -Descending) {
        try { $aForm.Controls.RemoveAt($index) } catch {}
    }
}

Function Set-Tooltip {
    param($Control, [string]$StrTooltip, [switch]$Show)
    $ToolTip = New-Object System.Windows.Forms.ToolTip
    $ToolTip.BackColor = [System.Drawing.Color]::LightGoldenrodYellow
    $ToolTip.IsBalloon = $true
    $ToolTip.InitialDelay = 500
    $ToolTip.ReshowDelay = 500
    $ToolTip.SetToolTip($Control, $StrTooltip)
    if ($Show) { $ToolTip.Show($StrTooltip, $Control, 1500) }
}

Function ASCinString($aString, $pos = 0) {
    $array = [char[]]"$($aString)" | % { [int]$_ }
    return $array[$pos]
}

Function CreatePassword($length = 12) {
    $find = @('ss','d','t','ee','a','i','j','J','z','e')
    $replace = @('$s',')','+','3e','@','!',']',']','2','3')
    $part1 = @()
    $part2 = @()
    [int]$lgth = (($length - 2) / 2)

    foreach ($name in $nouns) { if ($name.length -eq $lgth) { $part1 += $name } }
    foreach ($name in $verbs) { if ($name.length -eq $lgth) { $part2 += $name } }

    if ($part1.Count -eq 0 -or $part2.Count -eq 0) {
        # Fallback: genereer een veilig random wachtwoord
        $upper = 'ABCDEFGHJKLMNPQRSTUVWXYZ'
        $lower = 'abcdefghjkmnpqrstuvwxyz'
        $digits = '23456789'
        $special = '!@#$%&*?+-='
        $pwd = ''
        $pwd += $upper[(Get-Random -Maximum $upper.Length)]
        $pwd += $lower[(Get-Random -Maximum $lower.Length)]
        $pwd += $digits[(Get-Random -Maximum $digits.Length)]
        $pwd += $special[(Get-Random -Maximum $special.Length)]
        $all = $upper + $lower + $digits + $special
        for ($x = $pwd.Length; $x -lt $length; $x++) { $pwd += $all[(Get-Random -Maximum $all.Length)] }
        # Shuffle
        $pwd = -join ($pwd.ToCharArray() | Get-Random -Count $pwd.Length)
        return $pwd
    }

    do {
        $rnd1 = Get-Random -Minimum 1 -Maximum $part1.length
        $rnd2 = Get-Random -Minimum 1 -Maximum $part2.length
        $cap = Get-Random -Minimum 1 -Maximum 10
        $finalnums = Get-Random -Minimum 10 -Maximum 99
        $special = $false; $caps = $false

        $name1 = $part1[$rnd1 - 1]
        $name2 = $part2[$rnd2 - 1]

        if ($cap -gt 4) {
            $name1 = $name1.substring(0, 1).toupper() + $name1.substring(1).tolower()
            $ascii = ASCinString $name1 0
            if ($ascii -ge 65 -AND $ascii -le 90) { $caps = $true }
        } else {
            $name2 = $name2.substring(0, 1).toupper() + $name2.substring(1).tolower()
            $ascii = ASCinString $name2 0
            if ($ascii -ge 65 -AND $ascii -le 90) { $caps = $true }
        }

        $i = 0
        do {
            $f = $find[$i]; $r = $replace[$i]
            if ($name1 -clike "*$($f)*") {
                $regex = [regex]"$($f)"
                [string]$name1 = $regex.Replace($name1, "$($r)", 1)
                $special = $true
            } elseif ($name2 -clike "*$($f)*") {
                $regex = [regex]"$($f)"
                [string]$name2 = $regex.Replace($name2, "$($r)", 1)
                $special = $true
            }
            $i++
        } until ($special -eq $true -OR $i -ge $find.length)

        if ($caps -eq $false) {
            $ascii = ASCinString $name1 0
            if ($ascii -ge 97 -AND $ascii -le 122) {
                $name1 = $name1.substring(0, 1).toupper() + $name1.substring(1)
                $caps = $true
            } else {
                $ascii = ASCinString $name2 0
                if ($ascii -ge 97 -AND $ascii -le 122) {
                    $name2 = $name2.substring(0, 1).toupper() + $name2.substring(1)
                    $caps = $true
                }
            }
        }
    } until ($special -eq $true -AND $caps -eq $true)

    $result = "$($name1)$($name2)$($finalnums)"

    # Complexiteit garanderen: hoofdletter + kleine letter + cijfer + speciaal teken
    $hasUpper = $result -cmatch '[A-Z]'
    $hasLower = $result -cmatch '[a-z]'
    $hasDigit = $result -match '[0-9]'
    $hasSpecial = $result -match '[^a-zA-Z0-9]'
    if (-not $hasUpper) { $result = $result + 'A' }
    if (-not $hasLower) { $result = $result + 'x' }
    if (-not $hasDigit) { $result = $result + '7' }
    if (-not $hasSpecial) { $result = $result + '!' }

    return $result
}

Function PropperTitleCase($sometext) {
    $someText = (Get-Culture).TextInfo.ToTitleCase($someText)
    $someText = $someText.Replace("Macd","MacD").Replace("Vanh","VanH").Replace("Van Der ","van der ")
    $someText = $someText.Replace("De La ","de la ").Replace("Mcc","McC")
    $someText = $someText.Replace("O'f","O'F").Replace("O'k","O'K").Replace("O'd","O'D")
    $someText = $someText.Replace("O'b","O'B").Replace("O'c","O'C").Replace("-ann ","-Ann ")
    return $someText
}

Function GetYLoc($height, $Line = 1, $adjust = 0) {
    return [int](([int]$height * [int]$Line) - ([int]$height - 1) + 12) - [int]$adjust
}

Function GetXLoc($width, $Column = 1, $adjust = 0) {
    return [int](([int]$width * [int]$Column) - ([int]$width - 1) + 12) - [int]$adjust
}

Function Install-ModuleIfNotInstalled([string][Parameter(Mandatory = $true)]$moduleName, [string]$minimalVersion) {
    $module = Get-Module -Name $moduleName -ListAvailable | Where-Object { $null -eq $minimalVersion -or $minimalVersion -lt $_.Version } | Select-Object -Last 1
    if ($null -ne $module) {
        Write-Verbose ('Module {0} (v{1}) beschikbaar.' -f $moduleName, $module.Version)
    } else {
        Import-Module -Name 'PowershellGet'
        $installedModule = Get-InstalledModule -Name $moduleName -ErrorAction SilentlyContinue
        if ($null -eq $installedModule -or ($null -ne $minimalVersion -and $installedModule.Version -lt $minimalVersion)) {
            if ((Get-PackageProvider -Name NuGet -Force).Version -lt '2.8.5.201') {
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force
            }
            $optionalArgs = @{}
            if ($null -ne $minimalVersion) { $optionalArgs['RequiredVersion'] = $minimalVersion }
            Install-Module -Name $moduleName @optionalArgs -Scope CurrentUser -Force -Verbose
        }
    }
}

# ============================================================================
# RSAT Check (voor On-Prem modus)
# ============================================================================
Function Ensure-RSAT {
    $IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    $CheckAD = (Get-Module -Name ActiveDirectory -ListAvailable)

    if (-not $CheckAD -AND $IsAdmin -eq $False) {
        Start-Process powershell.exe "-sta -NoProfile -WindowStyle hidden -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        exit
    }
    if (-not $CheckAD -AND $IsAdmin -eq $True) {
        $install = (Get-WindowsCapability -Online | Where-Object { $_.Name -like "RSAT.Active*" -AND $_.State -eq "NotPresent" })
        foreach ($item in $install) {
            try { Add-WindowsCapability -Online -Name $item.name }
            catch {
                [System.Windows.MessageBox]::Show($_.Exception.Message, "RSAT Installatie", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
                exit
            }
        }
    }
    if (-not $CheckAD -AND $IsAdmin -eq $False) {
        [System.Windows.MessageBox]::Show("Administrator rechten nodig om RSAT te installeren", "RSAT", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        exit
    }
}

# ============================================================================
# Module installatie - wordt later aangeroepen op basis van omgevingskeuze
# ============================================================================
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Function Install-GraphModules {
    Write-Log "Graph modules controleren..."
    $requiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Users.Actions',
        'Microsoft.Graph.Groups',
        'Microsoft.Graph.Identity.DirectoryManagement'
    )

    # ── Stap 1: Check welke modules al geinstalleerd zijn ──
    $missingModules = @()
    foreach ($mod in $requiredModules) {
        $installed = Get-Module -Name $mod -ListAvailable -ErrorAction SilentlyContinue
        if ($installed) {
            Write-Log "  [OK] $mod v$($installed[0].Version) gevonden"
        } else {
            $missingModules += $mod
            Write-Log "  [--] $mod niet gevonden"
        }
    }

    # ── Stap 2: Installeer alleen ontbrekende modules ──
    if ($missingModules.Count -gt 0) {
        Write-Log "$($missingModules.Count) module(s) ontbreken, installatie starten..."
        try {
            # NuGet provider (eenmalig)
            if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue | Where-Object { $_.Version -ge '2.8.5.201' })) {
                Write-Log "NuGet provider installeren..."
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction Stop
                Write-Log "NuGet provider geinstalleerd" "SUCCESS"
            }

            # Bepaal scope: AllUsers als admin (server), CurrentUser als fallback
            $IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
            $installScope = if ($IsAdmin) { "AllUsers" } else { "CurrentUser" }
            Write-Log "Installatie scope: $installScope"

            foreach ($mod in $missingModules) {
                Write-Log "  Installeren: $mod..."
                Install-Module -Name $mod -Scope $installScope -Force -AllowClobber -SkipPublisherCheck -ErrorAction Stop
                Write-Log "  $mod geinstalleerd" "SUCCESS"
            }

            # PSModulePath verversen (nodig als modules net geinstalleerd zijn)
            $refreshPaths = @(
                "$env:ProgramFiles\WindowsPowerShell\Modules",
                "$env:ProgramFiles\PowerShell\Modules",
                (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'WindowsPowerShell\Modules'),
                (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PowerShell\Modules')
            )
            foreach ($p in $refreshPaths) {
                if ((Test-Path $p) -and $env:PSModulePath -notlike "*$p*") {
                    $env:PSModulePath = "$p;$env:PSModulePath"
                }
            }
        } catch {
            Write-Log "Module installatie mislukt: $($_.Exception.Message)" "ERROR"
            [System.Windows.Forms.MessageBox]::Show(
                "Microsoft Graph modules konden niet worden geinstalleerd.`n`n" +
                "Fout: $($_.Exception.Message)`n`n" +
                "Voer dit handmatig uit in een PowerShell venster:`n" +
                "Install-Module Microsoft.Graph -Scope AllUsers -Force`n`n" +
                "Start het script daarna opnieuw.",
                "Graph Modules Ontbreken",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return $false
        }
    } else {
        Write-Log "Alle Graph modules zijn al geinstalleerd" "SUCCESS"
    }

    # ── Stap 3: Modules laden ──
    $loadFailed = @()
    foreach ($mod in $requiredModules) {
        if (Get-Module -Name $mod) {
            # Al geladen in deze sessie, skip
            continue
        }
        try {
            Import-Module $mod -ErrorAction Stop -WarningAction SilentlyContinue
        } catch {
            Write-Log "  Import mislukt voor ${mod}: $($_.Exception.Message)" "WARN"
            # Fallback: zoek .psd1 rechtstreeks (cache probleem na verse installatie)
            $psd = Get-ChildItem -Path @("$env:ProgramFiles\WindowsPowerShell\Modules\$mod", "$env:ProgramFiles\PowerShell\Modules\$mod") -Filter "$mod.psd1" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($psd) {
                try {
                    Import-Module $psd.FullName -ErrorAction Stop -WarningAction SilentlyContinue
                    Write-Log "  $mod geladen via: $($psd.FullName)" "SUCCESS"
                } catch { $loadFailed += $mod }
            } else { $loadFailed += $mod }
        }
    }

    if ($loadFailed.Count -gt 0) {
        Write-Log "Modules niet geladen: $($loadFailed -join ', ')" "ERROR"
        Write-Log "Start het script opnieuw - na eerste installatie is een herstart soms nodig" "WARN"
        [System.Windows.Forms.MessageBox]::Show(
            "De volgende Graph modules konden niet worden geladen:`n$($loadFailed -join "`n")`n`n" +
            "Dit komt doordat de modules net zijn geinstalleerd.`n" +
            "Start het script opnieuw.",
            "Graph Modules - Herstart Nodig",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return $false
    }

    $loadedCount = ($requiredModules | Where-Object { Get-Module -Name $_ }).Count
    Write-Log "Graph modules geladen ($loadedCount/$($requiredModules.Count))" "SUCCESS"
    return $true
}

cls

# ============================================================================
# STAP 1: Omgevingskeuze formulier - On-Prem AD of Cloud Only
# ============================================================================
Function Show-EnvironmentSelector {

    $envForm = New-Object System.Windows.Forms.Form
    $envForm.Text = "Create 365 User - MSP Edition"
    $envForm.Size = New-Object System.Drawing.Size(450, 280)
    $envForm.StartPosition = "CenterScreen"
    $envForm.FormBorderStyle = 'FixedSingle'
    $envForm.MaximizeBox = $false
    $envForm.MinimizeBox = $false
    $envForm.BackColor = $BackColor
    $envForm.ForeColor = $ForeColor
    $envForm.Font = $Form_Font
    $envForm.TopMost = $true

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "Selecteer de omgeving"
    $lblTitle.Font = $Font14B
    $lblTitle.AutoSize = $true
    $lblTitle.Location = New-Object System.Drawing.Point(20, 20)
    [void]$envForm.Controls.Add($lblTitle)

    $lblDesc = New-Object System.Windows.Forms.Label
    $lblDesc.Text = "Kies het type omgeving voor de nieuwe gebruiker:"
    $lblDesc.AutoSize = $true
    $lblDesc.Location = New-Object System.Drawing.Point(20, 60)
    [void]$envForm.Controls.Add($lblDesc)

    $btnOnPrem = New-Object System.Windows.Forms.Button
    $btnOnPrem.Text = "On-Premises AD`n(met Azure AD Sync)"
    $btnOnPrem.Size = New-Object System.Drawing.Size(180, 80)
    $btnOnPrem.Location = New-Object System.Drawing.Point(20, 100)
    $btnOnPrem.FlatStyle = "Flat"
    $btnOnPrem.FlatAppearance.BorderColor = $ButtonBorderColor
    $btnOnPrem.FlatAppearance.BorderSize = 2
    $btnOnPrem.FlatAppearance.MouseDownBackColor = $ButtonMouseDownColor
    $btnOnPrem.FlatAppearance.MouseOverBackColor = $ButtonMouseOverColor
    $btnOnPrem.ForeColor = $ForeColor
    $btnOnPrem.BackColor = $ButtonColor
    $btnOnPrem.Font = $Font10B
    $btnOnPrem.Cursor = [System.Windows.Forms.Cursors]::Hand
    Set-Tooltip $btnOnPrem "Gebruiker wordt aangemaakt in Active Directory.`nAzure AD Sync repliceert naar Office 365."
    $btnOnPrem.Add_Click({
        $envForm.Tag = "OnPrem"
        $envForm.Close()
    })
    [void]$envForm.Controls.Add($btnOnPrem)

    $btnCloud = New-Object System.Windows.Forms.Button
    $btnCloud.Text = "Cloud Only`n(Office 365)"
    $btnCloud.Size = New-Object System.Drawing.Size(180, 80)
    $btnCloud.Location = New-Object System.Drawing.Point(230, 100)
    $btnCloud.FlatStyle = "Flat"
    $btnCloud.FlatAppearance.BorderColor = $ButtonBorderColor
    $btnCloud.FlatAppearance.BorderSize = 2
    $btnCloud.FlatAppearance.MouseDownBackColor = $ButtonMouseDownColor
    $btnCloud.FlatAppearance.MouseOverBackColor = $ButtonMouseOverColor
    $btnCloud.ForeColor = $ForeColor
    $btnCloud.BackColor = $ButtonColor
    $btnCloud.Font = $Font10B
    $btnCloud.Cursor = [System.Windows.Forms.Cursors]::Hand
    Set-Tooltip $btnCloud "Gebruiker wordt direct aangemaakt in Office 365.`nGeen Active Directory nodig."
    $btnCloud.Add_Click({
        $envForm.Tag = "CloudOnly"
        $envForm.Close()
    })
    [void]$envForm.Controls.Add($btnCloud)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuleren"
    $btnCancel.Size = New-Object System.Drawing.Size(100, 30)
    $btnCancel.Location = New-Object System.Drawing.Point(165, 200)
    $btnCancel.FlatStyle = "Flat"
    $btnCancel.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
    $btnCancel.FlatAppearance.BorderSize = 1
    $btnCancel.ForeColor = $ForeColor
    $btnCancel.BackColor = $BackColor
    $btnCancel.Font = $Font9
    $btnCancel.Add_Click({
        $envForm.Tag = $null
        $envForm.Close()
    })
    [void]$envForm.Controls.Add($btnCancel)

    [void]$envForm.ShowDialog()
    return $envForm.Tag
}

# ============================================================================
# STAP 2a: OU Selector formulier (On-Prem - draait lokaal op de DC)
# ============================================================================
Function Show-OUSelectorForm {

    # Haal OU's op uit AD
    $ouList = @()
    try {
        $ous = Get-ADOrganizationalUnit -Filter * -Properties CanonicalName -ErrorAction Stop | Sort-Object CanonicalName
        foreach ($ou in $ous) {
            $ouList += [PSCustomObject]@{
                Name = $ou.CanonicalName
                DN   = $ou.DistinguishedName
            }
        }
    } catch {
        Write-Log "OU's ophalen mislukt: $($_.Exception.Message)" "ERROR"
    }

    $ouForm = New-Object System.Windows.Forms.Form
    $ouForm.Text = "On-Premises AD - OU Selectie"
    $ouForm.Size = New-Object System.Drawing.Size(550, 250)
    $ouForm.StartPosition = "CenterScreen"
    $ouForm.FormBorderStyle = 'FixedSingle'
    $ouForm.MaximizeBox = $false
    $ouForm.MinimizeBox = $false
    $ouForm.BackColor = $BackColor
    $ouForm.ForeColor = $ForeColor
    $ouForm.Font = $Form_Font
    $ouForm.TopMost = $true

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "Selecteer de OU voor nieuwe gebruikers"
    $lblTitle.Font = $Font14B
    $lblTitle.AutoSize = $true
    $lblTitle.Location = New-Object System.Drawing.Point(20, 15)
    [void]$ouForm.Controls.Add($lblTitle)

    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = "Ingelogd als: $($env:USERDOMAIN)\$($env:USERNAME) op $($env:COMPUTERNAME)"
    $lblInfo.AutoSize = $true
    $lblInfo.Font = $Font8
    $lblInfo.ForeColor = [System.Drawing.Color]::Gray
    $lblInfo.Location = New-Object System.Drawing.Point(20, 50)
    [void]$ouForm.Controls.Add($lblInfo)

    $lblOU = New-Object System.Windows.Forms.Label
    $lblOU.Text = "OU:"
    $lblOU.AutoSize = $true
    $lblOU.Location = New-Object System.Drawing.Point(20, 85)
    [void]$ouForm.Controls.Add($lblOU)

    $cmbOU = New-Object System.Windows.Forms.ComboBox
    $cmbOU.Location = New-Object System.Drawing.Point(60, 83)
    $cmbOU.Width = 460
    $cmbOU.BackColor = $TextBoxBackColor
    $cmbOU.ForeColor = $ForeColor
    $cmbOU.FlatStyle = "Flat"
    $cmbOU.Font = $TextBoxFont
    $cmbOU.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    foreach ($ou in $ouList) { [void]$cmbOU.Items.Add($ou.Name) }
    # Probeer een default te selecteren met "Users" in de naam
    $defaultIdx = 0
    for ($i = 0; $i -lt $ouList.Count; $i++) {
        if ($ouList[$i].Name -like "*Users*" -or $ouList[$i].Name -like "*Gebruikers*") { $defaultIdx = $i; break }
    }
    if ($cmbOU.Items.Count -gt 0) { $cmbOU.SelectedIndex = $defaultIdx }
    Set-Tooltip $cmbOU "Organizational Unit waarin de nieuwe gebruiker wordt aangemaakt"
    [void]$ouForm.Controls.Add($cmbOU)

    # Checkbox: Azure AD Sync starten na aanmaken
    $chkSync = New-Object System.Windows.Forms.CheckBox
    $chkSync.Text = "Azure AD Sync starten na aanmaken (hybrid)"
    $chkSync.AutoSize = $true
    $chkSync.Location = New-Object System.Drawing.Point(60, 120)
    $chkSync.ForeColor = $ForeColor
    $chkSync.Font = $Font9
    $chkSync.Checked = $true
    Set-Tooltip $chkSync "Start een Delta Sync na het aanmaken en wacht tot de user in Office 365 verschijnt"
    [void]$ouForm.Controls.Add($chkSync)

    # Buttons
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "Doorgaan"
    $btnOK.Size = New-Object System.Drawing.Size(120, 30)
    $btnOK.Location = New-Object System.Drawing.Point(300, 165)
    $btnOK.FlatStyle = "Flat"
    $btnOK.FlatAppearance.BorderColor = $ButtonBorderColor
    $btnOK.FlatAppearance.BorderSize = 2
    $btnOK.FlatAppearance.MouseDownBackColor = $ButtonMouseDownColor
    $btnOK.FlatAppearance.MouseOverBackColor = $ButtonMouseOverColor
    $btnOK.ForeColor = $ForeColor
    $btnOK.BackColor = $ButtonColor
    $btnOK.Font = $Font10B
    $btnOK.Add_Click({
        if ($cmbOU.SelectedIndex -lt 0) {
            [System.Windows.Forms.MessageBox]::Show("Selecteer een OU.", "Fout", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
        $selectedOU = $ouList[$cmbOU.SelectedIndex]
        $ouForm.Tag = @{
            OUPath    = $selectedOU.DN
            OUName    = $selectedOU.Name
            StartSync = $chkSync.Checked
        }
        $ouForm.Close()
    })
    [void]$ouForm.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Annuleren"
    $btnCancel.Size = New-Object System.Drawing.Size(100, 30)
    $btnCancel.Location = New-Object System.Drawing.Point(150, 165)
    $btnCancel.FlatStyle = "Flat"
    $btnCancel.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray
    $btnCancel.FlatAppearance.BorderSize = 1
    $btnCancel.ForeColor = $ForeColor
    $btnCancel.BackColor = $BackColor
    $btnCancel.Font = $Font9
    $btnCancel.Add_Click({
        $ouForm.Tag = $null
        $ouForm.Close()
    })
    [void]$ouForm.Controls.Add($btnCancel)

    [void]$ouForm.ShowDialog()
    return $ouForm.Tag
}

# ============================================================================
# STAP 2a-2: Wacht tot user in 365 verschijnt na AD Sync
# ============================================================================
Function Start-ADSyncAndWait {
    param(
        [string]$UserPrincipalName,
        [string]$DisplayName,
        [int]$MaxWaitMinutes = 10,
        [int]$PollIntervalSeconds = 15
    )

    # Start Delta Sync
    Write-Log "Azure AD Sync starten (Delta)..."
    try {
        $syncResult = Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction Stop
        Write-Log "AD Sync gestart: $($syncResult.Result)" "SUCCESS"
    } catch {
        # Misschien draait ADSync niet op deze server, probeer via Invoke-Command
        Write-Log "Start-ADSyncSyncCycle niet beschikbaar op deze server, probeer remote..." "WARN"
        try {
            # Probeer de ADConnect server te vinden via MSOL service account
            $adcServer = (Get-ADUser -Filter "samAccountName -like 'MSOL_*'" -Properties Description -ErrorAction SilentlyContinue | Select-Object -First 1).Description
            if (-not $adcServer) {
                # Fallback: probeer via AADConnect service account
                $adcServer = (Get-ADUser -Filter "samAccountName -like 'AAD_*'" -Properties Description -ErrorAction SilentlyContinue | Select-Object -First 1).Description
            }
            if (-not $adcServer) {
                Write-Log "Kan ADConnect server niet automatisch vinden. Sync handmatig starten." "WARN"
                return $false
            }
            Invoke-Command -ComputerName $adcServer -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta } -ErrorAction Stop
            Write-Log "AD Sync remote gestart op $adcServer" "SUCCESS"
        } catch {
            Write-Log "AD Sync starten mislukt: $($_.Exception.Message)" "ERROR"
            Write-Log "Start de sync handmatig en wacht tot de user in 365 verschijnt." "WARN"
            return $false
        }
    }

    # Wacht tot user in 365 verschijnt
    Write-Log "Wachten tot $UserPrincipalName in Office 365 verschijnt (max $MaxWaitMinutes min)..."
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $maxMs = $MaxWaitMinutes * 60 * 1000
    $found = $false

    while ($stopwatch.ElapsedMilliseconds -lt $maxMs) {
        Start-Sleep -Seconds $PollIntervalSeconds
        $elapsed = [math]::Round($stopwatch.Elapsed.TotalSeconds)
        try {
            # Zoek op UPN
            $cloudUser = Get-MgUser -Filter "UserPrincipalName eq '$UserPrincipalName'" -ErrorAction SilentlyContinue
            # Fallback: zoek op mail
            if (-not $cloudUser) {
                $cloudUser = Get-MgUser -Filter "mail eq '$UserPrincipalName'" -ErrorAction SilentlyContinue
            }
            # Fallback: zoek op displayname
            if (-not $cloudUser -and $DisplayName) {
                $cloudUser = Get-MgUser -Filter "displayName eq '$DisplayName'" -ErrorAction SilentlyContinue | Select-Object -First 1
            }
            if ($cloudUser) {
                $found = $true
                Write-Log "User gevonden in Office 365 na $elapsed seconden: $($cloudUser.UserPrincipalName) (Id: $($cloudUser.Id))" "SUCCESS"
                break
            }
        } catch {}
        Write-Log "  Wachten... ($elapsed sec)" 
    }

    $stopwatch.Stop()
    if (-not $found) {
        Write-Log "User $UserPrincipalName niet gevonden in 365 na $MaxWaitMinutes minuten" "WARN"
    }
    return $found
}

# ============================================================================
# STAP 2b: MS Graph interactieve login + tenant info ophalen
# ============================================================================
Function Connect-GraphInteractive {
    param([string]$TenantId)
    $errorOccurred = $false
    try {
        $connectParams = @{
            Scopes = @("User.ReadWrite.All","Group.ReadWrite.All","Directory.ReadWrite.All","Organization.Read.All","Mail.Send","LicenseAssignment.ReadWrite.All")
            NoWelcome = $true
            ErrorAction = 'Stop'
        }
        if ($TenantId) {
            $connectParams['TenantId'] = $TenantId
            Write-Log "Graph login met TenantId: $TenantId"
        }
        # Debug en Warning output onderdrukken (voorkomt MSAL debug spam en WAM waarschuwing)
        $prevDebug = $DebugPreference; $prevWarn = $WarningPreference
        $DebugPreference = 'SilentlyContinue'; $WarningPreference = 'SilentlyContinue'
        Connect-MgGraph @connectParams
        $DebugPreference = $prevDebug; $WarningPreference = $prevWarn
        Start-Sleep -Seconds 1
    } catch {
        $e = $_.Exception
        Write-Log "Graph login exception: $($e.Message)" "ERROR"
        $errorOccurred = $true
    }

    # Verificatie: check of we echt verbonden zijn
    if (-not $errorOccurred) {
        $ctx = Get-MgContext -ErrorAction SilentlyContinue
        if (-not $ctx -or -not $ctx.Account) {
            Write-Log "Graph login leek te slagen maar Get-MgContext geeft geen account terug" "ERROR"
            $errorOccurred = $true
        }
    }

    if ($errorOccurred) {
        [System.Windows.Forms.MessageBox]::Show("Microsoft Graph login is mislukt of geannuleerd.`n`nProbeer opnieuw en sluit het login-venster niet voortijdig.", "Login Fout", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }

    $global:GraphConnected = $true
    return $true
}

Function Get-TenantInfo {
    # Haal organisatie info op
    try {
        $org = Get-MgOrganization -ErrorAction Stop
        if ($org) {
            $global:CompanyName = $org.DisplayName
            Write-Log "Organisatie opgehaald: $($global:CompanyName)"
        }
    } catch {
        Write-Log "Organisatie info ophalen mislukt: $($_.Exception.Message)" "WARN"
    }

    # Haal domeinen op
    try {
        $domains = Get-MgDomain -ErrorAction Stop
        $global:TenantDomains = @()
        foreach ($d in $domains) {
            if ($d.Id -notlike "*.onmicrosoft.com") {
                $global:TenantDomains += $d.Id
            }
            if ($d.IsDefault -eq $true) {
                $global:PrimaryDomain = $d.Id
            }
        }
        # Als primary domain een onmicrosoft is, pak dan de eerste custom
        if ($global:PrimaryDomain -like "*.onmicrosoft.com" -and $global:TenantDomains.Count -gt 0) {
            $global:PrimaryDomain = $global:TenantDomains[0]
        }
        # Voeg ook onmicrosoft domeinen toe
        foreach ($d in $domains) {
            if ($d.Id -like "*.onmicrosoft.com") {
                $global:TenantDomains += $d.Id
            }
        }
        Write-Log "Domeinen opgehaald: $($global:TenantDomains -join ', ')"
    } catch {
        Write-Log "Domeinen ophalen mislukt: $($_.Exception.Message)" "ERROR"
    }

    # Haal licenties op
    try {
        $global:TenantLicenses = Get-MgSubscribedSKU -All -Property @("SkuId","SkuPartNumber","ConsumedUnits","PrepaidUnits") -ErrorAction Stop |
            Select-Object *, @{Name = "ActiveUnits"; Expression = { ($_ | Select-Object -ExpandProperty PrepaidUnits).Enabled } } |
            Select-Object SkuId, SkuPartNumber, ActiveUnits, ConsumedUnits |
            Where-Object { $_.ActiveUnits -gt 0 }
        Write-Log "Licenties opgehaald: $($global:TenantLicenses.Count) typen"
    } catch {
        Write-Log "Licenties ophalen mislukt: $($_.Exception.Message)" "ERROR"
    }
}


# ============================================================================
# 365 Inrichtingsformulier (voor hybrid flow na sync)
# ============================================================================
Function Show-365ConfigForm {
    param(
        [string]$UserDisplayName,
        [string]$UserEmail,
        [string]$UserId
    )

    Write-Log "365 inrichtingsformulier laden..."
    $global:ExoConnected = $false

    # EXO verbinden (nodig voor shared mailboxes en distributielijsten)
    Write-Log "Exchange Online verbinden..."
    try {
        # OS detectie: Server 2019 (build 17763) heeft geen MSAL Broker
        # Server 2022 (build 20348+) en Server 2025 (build 26100+) wel
        $osBuild = [int](Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).BuildNumber
        $needsCompatible = ($osBuild -lt 20348)  # Server 2019 of ouder
        if ($needsCompatible) {
            Write-Log "Server 2019 of ouder gedetecteerd (build $osBuild) - EXO v3.2.0 vereist (geen MSAL Broker)"
        } else {
            Write-Log "Server 2022+ gedetecteerd (build $osBuild) - nieuwste EXO versie"
        }

        # Module installeren/laden
        if ($needsCompatible) {
            # Server 2019: gebruik v3.2.0 (laatste zonder broker dependency)
            $exoModule = Get-Module ExchangeOnlineManagement -ListAvailable -ErrorAction SilentlyContinue |
                Where-Object { $_.Version -le [Version]"3.2.0" } | Sort-Object Version -Descending | Select-Object -First 1
            if (-not $exoModule) {
                Write-Log "ExchangeOnlineManagement v3.2.0 installeren (Server 2019 compatibel)..."
                Install-Module ExchangeOnlineManagement -RequiredVersion 3.2.0 -Scope AllUsers -Force -AllowClobber -SkipPublisherCheck -ErrorAction Stop
                Write-Log "ExchangeOnlineManagement v3.2.0 geinstalleerd" "SUCCESS"
            } else {
                Write-Log "ExchangeOnlineManagement v$($exoModule.Version) gevonden (compatibel)"
            }
            # Zorg dat v3.2.0 geladen wordt (niet een eventuele nieuwere versie)
            Remove-Module ExchangeOnlineManagement -Force -ErrorAction SilentlyContinue
            Import-Module ExchangeOnlineManagement -MaximumVersion 3.2.0 -ErrorAction Stop -WarningAction SilentlyContinue
        } else {
            # Server 2022/2025: gebruik nieuwste versie
            $exoModule = Get-Module ExchangeOnlineManagement -ListAvailable -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1
            if (-not $exoModule) {
                Write-Log "ExchangeOnlineManagement module installeren..."
                Install-Module ExchangeOnlineManagement -Scope AllUsers -Force -AllowClobber -SkipPublisherCheck -ErrorAction Stop
                $exoModule = Get-Module ExchangeOnlineManagement -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
                Write-Log "ExchangeOnlineManagement v$($exoModule.Version) geinstalleerd" "SUCCESS"
            } else {
                Write-Log "ExchangeOnlineManagement v$($exoModule.Version) gevonden"
            }
            Import-Module ExchangeOnlineManagement -ErrorAction Stop -WarningAction SilentlyContinue
        }

        $ctx = Get-MgContext
        if ($ctx) {
            $exoConnected = $false
            $prevDebug = $DebugPreference; $prevWarn = $WarningPreference
            $DebugPreference = 'SilentlyContinue'; $WarningPreference = 'SilentlyContinue'

            # Poging 1: Login met UPN (werkt op alle versies)
            try {
                Connect-ExchangeOnline -UserPrincipalName $ctx.Account -ShowBanner:$false -ErrorAction Stop
                $exoConnected = $true
                Write-Log "Exchange Online verbonden" "SUCCESS"
            } catch {
                Write-Log "EXO poging 1 (UPN) mislukt: $($_.Exception.Message)" "WARN"
            }

            # Poging 2: Interactief zonder UPN
            if (-not $exoConnected) {
                try {
                    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
                    $exoConnected = $true
                    Write-Log "Exchange Online verbonden (interactief)" "SUCCESS"
                } catch {
                    Write-Log "EXO poging 2 (interactief) mislukt: $($_.Exception.Message)" "WARN"
                }
            }

            $DebugPreference = $prevDebug; $WarningPreference = $prevWarn
            $global:ExoConnected = $exoConnected
            if (-not $exoConnected) {
                Write-Log "Exchange Online kon niet verbinden - DLs en shared mailboxes niet beschikbaar" "WARN"
            }
        }
    } catch {
        Write-Log "Exchange Online module laden mislukt: $($_.Exception.Message)" "WARN"
    }

    # Cloud groepen ophalen via Graph
    Write-Log "Cloud groepen ophalen..."
    $allGroups = @()
    try {
        $rawGroups = Get-MgGroup -All -Property 'Id,DisplayName,GroupTypes,SecurityEnabled,MailEnabled,Mail,OnPremisesSyncEnabled' -ErrorAction Stop
        $allGroups = $rawGroups | Where-Object { $_.OnPremisesSyncEnabled -ne $true } | Sort-Object DisplayName
        Write-Log "Cloud groepen opgehaald: $($allGroups.Count)"
    } catch { Write-Log "Groepen ophalen mislukt: $($_.Exception.Message)" "ERROR" }

    # Categoriseer
    $m365Groups = @(); $distLists = @(); $secGroups = @()
    foreach ($g in $allGroups) {
        $types = $g.GroupTypes -join ","
        if ($types -like "*Unified*") { $m365Groups += $g }
        elseif ($g.MailEnabled -eq $true -and $g.SecurityEnabled -ne $true) { $distLists += $g }
        else { $secGroups += $g }
    }

    # Shared mailboxes ALLEEN via EXO (Graph is onbetrouwbaar hiervoor)
    $sharedMailboxes = @()
    if ($global:ExoConnected) {
        Write-Log "Shared mailboxes ophalen via Exchange Online..."
        try {
            $sharedMailboxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop |
                Select-Object DisplayName, PrimarySmtpAddress, ExternalDirectoryObjectId | Sort-Object DisplayName
            Write-Log "Shared mailboxes: $($sharedMailboxes.Count)" "SUCCESS"
        } catch { Write-Log "Shared mailboxes ophalen mislukt: $($_.Exception.Message)" "WARN" }

        # Extra DLs via EXO
        try {
            $exoDLs = Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop |
                Select-Object DisplayName, PrimarySmtpAddress, ExternalDirectoryObjectId | Sort-Object DisplayName
            $existingDLNames = $distLists | ForEach-Object { $_.DisplayName }
            $newDLs = 0
            foreach ($dl in $exoDLs) {
                if ($dl.DisplayName -notin $existingDLNames) {
                    $distLists += [PSCustomObject]@{
                        Id = $dl.ExternalDirectoryObjectId; DisplayName = $dl.DisplayName
                        Mail = $dl.PrimarySmtpAddress; MailEnabled = $true
                        SecurityEnabled = $false; GroupTypes = @(); OnPremisesSyncEnabled = $false
                    }
                    $newDLs++
                }
            }
            if ($newDLs -gt 0) { Write-Log "  $newDLs extra DLs via EXO (totaal: $($distLists.Count))" }
        } catch { Write-Log "EXO DLs ophalen mislukt: $($_.Exception.Message)" "WARN" }
    } else {
        Write-Log "EXO niet verbonden - shared mailboxes en extra DLs niet beschikbaar" "WARN"
    }

    Write-Log "  M365: $($m365Groups.Count) | DLs: $($distLists.Count) | Security: $($secGroups.Count) | Shared MB: $($sharedMailboxes.Count)"

    # === FORMULIER BOUWEN ===
    $cfgForm = New-Object System.Windows.Forms.Form
    $cfgForm.Text = "365 Inrichting - $UserDisplayName ($UserEmail)"
    $cfgForm.Size = New-Object System.Drawing.Size(700, 630)
    $cfgForm.StartPosition = "CenterScreen"; $cfgForm.FormBorderStyle = 'FixedSingle'
    $cfgForm.MaximizeBox = $false; $cfgForm.TopMost = $true
    $cfgForm.BackColor = $BackColor; $cfgForm.ForeColor = $ForeColor; $cfgForm.Font = $Form_Font

    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(15, 15)
    $tabControl.Size = New-Object System.Drawing.Size(655, 500)
    $tabControl.Font = $Form_Font
    [void]$cfgForm.Controls.Add($tabControl)

    # === Helper: tab met zoekfunctie ===
    Function New-SearchableTab {
        param($TabControl, $Title, $Items, $KeyProperty, $DisplayProperty, $ExtraProperty)
        $tab = New-Object System.Windows.Forms.TabPage
        $tab.Text = $Title; $tab.BackColor = $BackColor; $tab.ForeColor = $ForeColor

        $searchBox = New-Object System.Windows.Forms.TextBox
        $searchBox.Location = New-Object System.Drawing.Point(10, 12); $searchBox.Width = 520
        $searchBox.BackColor = $TextBoxBackColor; $searchBox.ForeColor = [System.Drawing.Color]::Gray
        $searchBox.Font = $TextBoxFont; $searchBox.Text = "Zoek..."
        $searchBox.Add_GotFocus({ if ($this.Text -eq "Zoek...") { $this.Text = ""; $this.ForeColor = $ForeColor } })
        $searchBox.Add_LostFocus({ if ([string]::IsNullOrWhiteSpace($this.Text)) { $this.Text = "Zoek..."; $this.ForeColor = [System.Drawing.Color]::Gray } })
        [void]$tab.Controls.Add($searchBox)

        $countLabel = New-Object System.Windows.Forms.Label
        $countLabel.Text = "$($Items.Count) items"; $countLabel.AutoSize = $true
        $countLabel.Location = New-Object System.Drawing.Point(540, 15)
        $countLabel.ForeColor = [System.Drawing.Color]::Gray; $countLabel.Font = $Font8
        [void]$tab.Controls.Add($countLabel)

        $listBox = New-Object System.Windows.Forms.CheckedListBox
        $listBox.Location = New-Object System.Drawing.Point(10, 40); $listBox.Size = New-Object System.Drawing.Size(620, 420)
        $listBox.BackColor = $TextBoxBackColor; $listBox.ForeColor = $ForeColor
        $listBox.Font = $TextBoxFont; $listBox.CheckOnClick = $true
        $listBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        [void]$tab.Controls.Add($listBox)

        $checkedSet = [System.Collections.Generic.HashSet[string]]::new()
        $tab.Tag = @{ Items = $Items; ListBox = $listBox; CheckedSet = $checkedSet; KeyProperty = $KeyProperty; DisplayProperty = $DisplayProperty; ExtraProperty = $ExtraProperty }

        foreach ($item in $Items) {
            $display = $item.$DisplayProperty
            if ($ExtraProperty -and $item.$ExtraProperty) { $display = "$display  [$($item.$ExtraProperty)]" }
            [void]$listBox.Items.Add($display)
        }

        $searchBox.Add_TextChanged({
            $searchText = $this.Text
            if ($searchText -eq "Zoek...") { return }
            $tagData = $this.Parent.Tag; $lb = $tagData.ListBox; $cs = $tagData.CheckedSet
            for ($i = 0; $i -lt $lb.Items.Count; $i++) {
                $iName = $lb.Items[$i].ToString()
                if ($lb.GetItemChecked($i)) { [void]$cs.Add($iName) } else { [void]$cs.Remove($iName) }
            }
            $lb.Items.Clear()
            foreach ($item in $tagData.Items) {
                $display = $item.($tagData.DisplayProperty)
                if ($tagData.ExtraProperty -and $item.($tagData.ExtraProperty)) { $display = "$display  [$($item.($tagData.ExtraProperty))]" }
                if ([string]::IsNullOrWhiteSpace($searchText) -or $display -like "*$searchText*") {
                    [void]$lb.Items.Add($display)
                    if ($cs.Contains($display)) { $lb.SetItemChecked($lb.Items.Count - 1, $true) }
                }
            }
        }.GetNewClosure())

        [void]$TabControl.TabPages.Add($tab)
        return $tab
    }

    # TAB 1: 365 Groepen
    $grpItems = @()
    foreach ($g in $m365Groups) { $grpItems += [PSCustomObject]@{ Id = $g.Id; Name = $g.DisplayName; Extra = "M365"; Mail = $g.Mail } }
    foreach ($g in $secGroups) { $grpItems += [PSCustomObject]@{ Id = $g.Id; Name = $g.DisplayName; Extra = "Security"; Mail = $g.Mail } }
    $tab1 = New-SearchableTab -TabControl $tabControl -Title "365 Groepen ($($grpItems.Count))" -Items $grpItems -KeyProperty 'Id' -DisplayProperty 'Name' -ExtraProperty 'Extra'

    # TAB 2: Distributielijsten
    $dlItems = @()
    foreach ($g in $distLists) {
        $mail = if ($g.Mail) { $g.Mail } elseif ($g.PrimarySmtpAddress) { $g.PrimarySmtpAddress } else { "" }
        $dlItems += [PSCustomObject]@{ Id = $g.Id; Name = $g.DisplayName; Extra = $mail; Mail = $mail }
    }
    $tab2 = New-SearchableTab -TabControl $tabControl -Title "Distributielijsten ($($dlItems.Count))" -Items $dlItems -KeyProperty 'Id' -DisplayProperty 'Name' -ExtraProperty 'Extra'

    # TAB 3: Shared Mailboxes (alleen via EXO)
    $mbItems = @()
    foreach ($mb in $sharedMailboxes) { $mbItems += [PSCustomObject]@{ Id = $mb.ExternalDirectoryObjectId; Name = $mb.DisplayName; Extra = $mb.PrimarySmtpAddress; Mail = $mb.PrimarySmtpAddress } }
    $tab3 = New-SearchableTab -TabControl $tabControl -Title "Shared Mailboxes ($($mbItems.Count))" -Items $mbItems -KeyProperty 'Id' -DisplayProperty 'Name' -ExtraProperty 'Extra'

    if (-not $global:ExoConnected) {
        $lblExo = New-Object System.Windows.Forms.Label
        $lblExo.Text = "Exchange Online niet verbonden - DLs en shared mailboxes niet beschikbaar.`nInstalleer: Install-Module ExchangeOnlineManagement -Scope AllUsers"
        $lblExo.ForeColor = [System.Drawing.Color]::Orange; $lblExo.AutoSize = $true
        $lblExo.Location = New-Object System.Drawing.Point(20, 50)
        [void]$tab3.Controls.Add($lblExo); $lblExo.BringToFront()
    }

    # Info label
    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Text = "Selecteer items en klik Toepassen. Bij DLs/shared mailboxes wordt gewacht tot de mailbox actief is."
    $lblInfo.AutoSize = $true; $lblInfo.Location = New-Object System.Drawing.Point(15, 522)
    $lblInfo.ForeColor = [System.Drawing.Color]::Gray; $lblInfo.Font = $Font8
    [void]$cfgForm.Controls.Add($lblInfo)

    # Knoppen
    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Text = "Toepassen"; $btnApply.Size = New-Object System.Drawing.Size(130, 32)
    $btnApply.Location = New-Object System.Drawing.Point(540, 548)
    $btnApply.FlatStyle = "Flat"; $btnApply.BackColor = $ButtonColor; $btnApply.ForeColor = $ForeColor; $btnApply.Font = $Font10B
    $btnApply.FlatAppearance.BorderColor = $ButtonBorderColor; $btnApply.FlatAppearance.BorderSize = 2
    $btnApply.Add_Click({ $cfgForm.DialogResult = [System.Windows.Forms.DialogResult]::OK; $cfgForm.Close() })
    [void]$cfgForm.Controls.Add($btnApply)

    $btnSkip = New-Object System.Windows.Forms.Button
    $btnSkip.Text = "Overslaan"; $btnSkip.Size = New-Object System.Drawing.Size(100, 32)
    $btnSkip.Location = New-Object System.Drawing.Point(420, 548)
    $btnSkip.FlatStyle = "Flat"; $btnSkip.BackColor = $BackColor; $btnSkip.ForeColor = $ForeColor; $btnSkip.Font = $Font9
    $btnSkip.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray; $btnSkip.FlatAppearance.BorderSize = 1
    $btnSkip.Add_Click({ $cfgForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $cfgForm.Close() })
    [void]$cfgForm.Controls.Add($btnSkip)

    $result = $cfgForm.ShowDialog()

    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Log "365 inrichting overgeslagen door gebruiker"
        $cfgForm.Dispose(); return $null
    }

    # Verzamel selecties
    $selections = @{ Groups = @(); DistLists = @(); SharedMailboxes = @(); ExoAvailable = $global:ExoConnected }
    foreach ($tabPage in @($tab1, $tab2, $tab3)) {
        $tagData = $tabPage.Tag; $lb = $tagData.ListBox; $cs = $tagData.CheckedSet
        for ($i = 0; $i -lt $lb.Items.Count; $i++) {
            if ($lb.GetItemChecked($i)) { [void]$cs.Add($lb.Items[$i].ToString()) }
        }
    }
    foreach ($name in $tab1.Tag.CheckedSet) {
        $cleanName = ($name -split "  \[")[0].Trim()
        $match = $grpItems | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
        if ($match) { $selections.Groups += $match }
    }
    foreach ($name in $tab2.Tag.CheckedSet) {
        $cleanName = ($name -split "  \[")[0].Trim()
        $match = $dlItems | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
        if ($match) { $selections.DistLists += $match }
    }
    foreach ($name in $tab3.Tag.CheckedSet) {
        $cleanName = ($name -split "  \[")[0].Trim()
        $match = $mbItems | Where-Object { $_.Name -eq $cleanName } | Select-Object -First 1
        if ($match) { $selections.SharedMailboxes += $match }
    }

    Write-Log "Selecties: $($selections.Groups.Count) groepen, $($selections.DistLists.Count) DLs, $($selections.SharedMailboxes.Count) shared mailboxes"
    $cfgForm.Dispose()
    return $selections
}

# ============================================================================
# Wacht tot mailbox is aangemaakt na licentie toewijzing
# ============================================================================
Function Wait-ForMailbox {
    param(
        [string]$UserEmail,
        [int]$MaxWaitMinutes = 5,
        [int]$PollIntervalSeconds = 15
    )

    Write-Log "Wachten tot mailbox beschikbaar is voor $UserEmail (max $MaxWaitMinutes min)..."
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $maxMs = $MaxWaitMinutes * 60 * 1000

    while ($stopwatch.ElapsedMilliseconds -lt $maxMs) {
        try {
            $mbx = Get-EXOMailbox -Identity $UserEmail -ErrorAction Stop
            if ($mbx) {
                $elapsed = [math]::Round($stopwatch.Elapsed.TotalSeconds)
                Write-Log "Mailbox gevonden voor $UserEmail na $elapsed seconden" "SUCCESS"
                return $true
            }
        } catch {
            # Mailbox bestaat nog niet - normaal gedrag
        }
        $remaining = [math]::Round(($maxMs - $stopwatch.ElapsedMilliseconds) / 1000)
        Write-Host "  Wachten op mailbox... ($remaining sec resterend)" -ForegroundColor Gray
        Start-Sleep -Seconds $PollIntervalSeconds
    }

    Write-Log "Mailbox niet gevonden na $MaxWaitMinutes minuten" "WARN"
    return $false
}

# ============================================================================
# 365 Selecties verwerken (groepen direct, DLs/mailboxes na mailbox check)
# ============================================================================
Function Process-365Selections {
    param($Selections, [string]$UserId, [string]$UserEmail, [bool]$ExoAvailable, [bool]$LicenseAssigned)

    if (-not $Selections) { return }
    $totalOK = 0; $totalFailed = 0; $totalSkipped = 0

    # STAP 1: 365 Groepen toevoegen (geen mailbox nodig)
    if ($Selections.Groups.Count -gt 0) {
        Write-Log "=== 365 Groepen toevoegen ($($Selections.Groups.Count)) ==="
        foreach ($grp in $Selections.Groups) {
            try {
                New-MgGroupMember -GroupId $grp.Id -DirectoryObjectId $UserId -ErrorAction Stop | Out-Null
                Write-Log "  [OK] $($grp.Name)" "SUCCESS"; $totalOK++
            } catch {
                if ($_.Exception.Message -like "*already exist*") {
                    Write-Log "  [SKIP] $($grp.Name) - al lid" "WARN"; $totalSkipped++
                } else {
                    Write-Log "  [FAIL] $($grp.Name): $($_.Exception.Message)" "ERROR"; $totalFailed++
                }
            }
        }
    }

    # STAP 2: DLs en Shared Mailboxes vereisen een actieve mailbox
    $needsMailbox = ($Selections.DistLists.Count -gt 0) -or ($Selections.SharedMailboxes.Count -gt 0)

    if ($needsMailbox) {
        if (-not $LicenseAssigned) {
            Write-Log "Geen licentie toegewezen - DLs en shared mailboxes overgeslagen (mailbox vereist)" "WARN"
            Write-Log "Wijs eerst een licentie toe en voeg DLs/shared mailboxes handmatig toe" "WARN"
            return @{ OK = $totalOK; Failed = $totalFailed; Skipped = $totalSkipped }
        }

        if (-not $ExoAvailable) {
            Write-Log "Exchange Online niet verbonden - DLs en shared mailboxes overgeslagen" "WARN"
            return @{ OK = $totalOK; Failed = $totalFailed; Skipped = $totalSkipped }
        }

        # Wacht tot mailbox is aangemaakt
        Write-Log "Licentie is toegewezen, wachten tot mailbox wordt aangemaakt..."
        $mailboxReady = Wait-ForMailbox -UserEmail $UserEmail -MaxWaitMinutes 5 -PollIntervalSeconds 15

        if (-not $mailboxReady) {
            Write-Log "Mailbox nog niet beschikbaar - DLs en shared mailboxes overgeslagen" "WARN"
            Write-Log "Voeg DLs en shared mailboxes handmatig toe zodra de mailbox actief is" "WARN"
            return @{ OK = $totalOK; Failed = $totalFailed; Skipped = $totalSkipped }
        }

        # STAP 2a: Distributielijsten
        if ($Selections.DistLists.Count -gt 0) {
            Write-Log "=== Distributielijsten toevoegen ($($Selections.DistLists.Count)) ==="
            foreach ($dl in $Selections.DistLists) {
                try {
                    Add-DistributionGroupMember -Identity $dl.Mail -Member $UserEmail -ErrorAction Stop
                    Write-Log "  [OK] $($dl.Name)" "SUCCESS"; $totalOK++
                } catch {
                    if ($_.Exception.Message -like "*already a member*") {
                        Write-Log "  [SKIP] $($dl.Name) - al lid" "WARN"; $totalSkipped++
                    } else {
                        # Fallback via Graph
                        try {
                            New-MgGroupMember -GroupId $dl.Id -DirectoryObjectId $UserId -ErrorAction Stop | Out-Null
                            Write-Log "  [OK] $($dl.Name) (via Graph)" "SUCCESS"; $totalOK++
                        } catch {
                            if ($_.Exception.Message -like "*already exist*") {
                                Write-Log "  [SKIP] $($dl.Name) - al lid" "WARN"; $totalSkipped++
                            } else {
                                Write-Log "  [FAIL] $($dl.Name): $($_.Exception.Message)" "ERROR"; $totalFailed++
                            }
                        }
                    }
                }
            }
        }

        # STAP 2b: Shared Mailbox permissies
        if ($Selections.SharedMailboxes.Count -gt 0) {
            Write-Log "=== Shared Mailbox permissies ($($Selections.SharedMailboxes.Count)) ==="
            foreach ($mb in $Selections.SharedMailboxes) {
                # FullAccess + AutoMapping
                try {
                    Add-MailboxPermission -Identity $mb.Mail -User $UserEmail -AccessRights FullAccess -AutoMapping $true -ErrorAction Stop | Out-Null
                    Write-Log "  [OK] $($mb.Name) - FullAccess + AutoMapping" "SUCCESS"; $totalOK++
                } catch {
                    if ($_.Exception.Message -like "*already*") {
                        Write-Log "  [SKIP] $($mb.Name) FullAccess - bestaat al" "WARN"; $totalSkipped++
                    } else {
                        Write-Log "  [FAIL] $($mb.Name) FullAccess: $($_.Exception.Message)" "ERROR"; $totalFailed++
                    }
                }
                # SendAs
                try {
                    Add-RecipientPermission -Identity $mb.Mail -Trustee $UserEmail -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                    Write-Log "  [OK] $($mb.Name) - SendAs" "SUCCESS"
                } catch {
                    if ($_.Exception.Message -like "*already*") {
                        Write-Log "  [INFO] $($mb.Name) SendAs - bestaat al"
                    } else {
                        Write-Log "  [WARN] $($mb.Name) SendAs: $($_.Exception.Message)" "WARN"
                    }
                }
            }
        }
    }

    Write-Log "=== 365 Inrichting totaal: $totalOK OK, $totalFailed mislukt, $totalSkipped overgeslagen ==="
    return @{ OK = $totalOK; Failed = $totalFailed; Skipped = $totalSkipped }
}

Function Get-ADUsersLocal {
    param([string]$OUPath)
    $users = @()
    try {
        if ($OUPath) {
            $users = Get-ADUser -SearchBase $OUPath -Filter "enabled -eq 'true'" -Properties Name, Department, Title, SamAccountName, UserPrincipalName, DisplayName, DistinguishedName, MemberOf
        } else {
            $users = Get-ADUser -Filter "enabled -eq 'true'" -Properties Name, Department, Title, SamAccountName, UserPrincipalName, DisplayName, DistinguishedName, MemberOf
        }
    } catch {
        Write-Log "AD users ophalen mislukt: $($_.Exception.Message)" "ERROR"
    }
    return $users
}

# ============================================================================
# AD Group helper (lokaal)
# ============================================================================
Function AddUserToADGroup {
    param([string]$GroupName, [string]$UserSamAccountName)
    try {
        $grp = Get-ADGroup -LDAPFilter "(SAMAccountName=$GroupName)" -ErrorAction SilentlyContinue
        if (-not $grp) { return "AD Group '$GroupName' bestaat niet." }
        $existing = Get-ADGroupMember -Identity $GroupName -ErrorAction SilentlyContinue | Where-Object { $_.SamAccountName -eq $UserSamAccountName }
        if ($existing) { return "ALREADY_MEMBER" }
        Add-ADGroupMember -Identity $GroupName -Members $UserSamAccountName -ErrorAction Stop
        return $True
    } catch {
        return $_.Exception.Message
    }
}

# ============================================================================
# HOOFDPROGRAMMA - Start hier
# ============================================================================

# --- STAP 1: Omgevingskeuze ---
$global:EnvironmentMode = Show-EnvironmentSelector
if (-not $global:EnvironmentMode) {
    Write-Host "Geannuleerd."
    exit
}

Write-Host "Omgeving: $($global:EnvironmentMode)" -ForegroundColor Cyan
Write-Log "Script gestart - Omgeving: $($global:EnvironmentMode)"

# --- STAP 2a: On-Prem - OU selectie + lokale AD ---
$OUPath = ""
$global:StartSync = $false
if ($global:EnvironmentMode -eq "OnPrem") {
    # Admin rechten check - herstart als admin indien nodig
    $IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
    if (-not $IsAdmin) {
        Write-Log "Niet als Administrator gestart - herstarten met elevatie..." "WARN"
        try {
            Start-Process powershell.exe "-sta -NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Dit script moet als Administrator worden uitgevoerd voor On-Prem modus.`n`nKlik met rechtermuisknop op PowerShell en kies 'Als administrator uitvoeren'.",
                "Administrator Vereist", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning
            )
        }
        exit
    }
    # RSAT en ActiveDirectory module installeren indien nodig
    $adModule = Get-Module -Name ActiveDirectory -ListAvailable
    if (-not $adModule) {
        Write-Log "ActiveDirectory module niet gevonden, RSAT installeren..." "WARN"
        $IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
        if (-not $IsAdmin) {
            Write-Log "Administrator rechten nodig - script herstarten als admin..."
            Start-Process powershell.exe "-sta -NoProfile -WindowStyle hidden -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
            exit
        }
        try {
            $rsatFeatures = Get-WindowsCapability -Online | Where-Object { $_.Name -like "Rsat.ActiveDirectory*" -and $_.State -eq "NotPresent" }
            foreach ($feat in $rsatFeatures) {
                Write-Log "  Installeren: $($feat.Name)..."
                Add-WindowsCapability -Online -Name $feat.Name -ErrorAction Stop
                Write-Log "  Geinstalleerd" "SUCCESS"
            }
        } catch {
            Write-Log "RSAT installatie mislukt: $($_.Exception.Message)" "ERROR"
            [System.Windows.Forms.MessageBox]::Show(
                "ActiveDirectory module kon niet worden geinstalleerd.`n`nFout: $($_.Exception.Message)`n`nInstalleer RSAT handmatig via Windows Instellingen > Apps > Optionele onderdelen.",
                "Module Ontbreekt", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error
            )
            exit
        }
    }
    try {
        Import-Module 'ActiveDirectory' -ErrorAction Stop -WarningAction SilentlyContinue
        Write-Log "ActiveDirectory module geladen" "SUCCESS"
    } catch {
        Write-Log "ActiveDirectory module laden mislukt: $($_.Exception.Message)" "ERROR"
        [System.Windows.Forms.MessageBox]::Show("ActiveDirectory module kan niet worden geladen.`n$($_.Exception.Message)", "Module Fout", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit
    }

    Write-Log "Ingelogd als: $($env:USERDOMAIN)\$($env:USERNAME) op $($env:COMPUTERNAME)"

    $ouConfig = Show-OUSelectorForm
    if (-not $ouConfig) {
        Write-Log "Geannuleerd door gebruiker bij OU selectie" "WARN"
        exit
    }

    $OUPath = $ouConfig.OUPath
    $global:StartSync = $ouConfig.StartSync
    Write-Log "OU geselecteerd: $($ouConfig.OUName) ($OUPath)"
    Write-Log "Azure AD Sync na aanmaken: $($global:StartSync)"
}

# --- STAP 2b: Graph Login (alleen bij Cloud Only) ---
if ($global:EnvironmentMode -eq "CloudOnly") {
    $modulesOK = Install-GraphModules
    if (-not $modulesOK) { Write-Log "Graph modules niet beschikbaar - script stopt" "ERROR"; exit }
    $graphOK = Connect-GraphInteractive
    if (-not $graphOK) { Write-Log "Graph login mislukt of geannuleerd" "ERROR"; exit }
    Write-Log "Microsoft Graph verbonden als: $((Get-MgContext).Account)" "SUCCESS"
}

# --- STAP 3: Tenant info ophalen ---
if ($global:EnvironmentMode -eq "CloudOnly") {
    Write-Log "Tenant informatie ophalen..."
    Get-TenantInfo
    Write-Log "Tenant: $($global:CompanyName) | Primair domein: $($global:PrimaryDomain) | Domeinen: $($global:TenantDomains -join ', ')"
    Write-Log "Licenties gevonden: $($global:TenantLicenses.Count) typen"
} else {
    # On-Prem: haal domein info + UPN suffixen uit AD
    try {
        $adDomain = Get-ADDomain -ErrorAction SilentlyContinue
        $global:CompanyName = $adDomain.Forest
        $global:PrimaryDomain = $adDomain.DNSRoot
        $global:TenantDomains = @($adDomain.DNSRoot)

        # UPN suffixen ophalen (extra domeinen zoals eptummers.nl)
        try {
            $adForest = Get-ADForest -ErrorAction SilentlyContinue
            if ($adForest.UPNSuffixes) {
                foreach ($suffix in $adForest.UPNSuffixes) {
                    if ($suffix -notin $global:TenantDomains) {
                        $global:TenantDomains += $suffix
                    }
                }
                # Eerste UPN suffix als primary als het een .local domein is
                if ($global:PrimaryDomain -like "*.local" -and $adForest.UPNSuffixes.Count -gt 0) {
                    $global:PrimaryDomain = $adForest.UPNSuffixes[0]
                }
            }
        } catch {
            Write-Log "UPN suffixen ophalen mislukt: $($_.Exception.Message)" "WARN"
        }

        Write-Log "AD Domein: $($adDomain.DNSRoot) | UPN Suffixen: $($global:TenantDomains -join ', ')"
    } catch {
        Write-Log "AD domein info ophalen mislukt: $($_.Exception.Message)" "WARN"
    }
}

# --- STAP 4: Users ophalen ---
$global:ADusers = @()
$adNames = @()
$adDepartments = @()
$adTitles = @()

if ($global:EnvironmentMode -eq "OnPrem") {
    Write-Log "AD gebruikers ophalen (lokaal)..."
    $global:ADusers = Get-ADUsersLocal -OUPath $OUPath
    $adNames = ($global:ADusers | Select-Object -ExpandProperty Name | Sort-Object)
    $adDepartments = ($global:ADusers | Where-Object { $_.Department } | Select-Object -ExpandProperty Department -Unique | Sort-Object)
    $adTitles = ($global:ADusers | Where-Object { $_.Title } | Select-Object -ExpandProperty Title -Unique | Sort-Object)
    Write-Log "AD: $($global:ADusers.Count) gebruikers, $($adDepartments.Count) afdelingen, $($adTitles.Count) functies opgehaald"
} else {
    Write-Log "365 gebruikers ophalen..."
    try {
        $graphUsers = Get-MgUser -All -Property DisplayName, UserPrincipalName, Department, JobTitle, Id
        $global:ADusers = $graphUsers
        $adNames = ($graphUsers | Select-Object -ExpandProperty DisplayName | Sort-Object)
        $adDepartments = ($graphUsers | Where-Object { $_.Department } | Select-Object -ExpandProperty Department -Unique | Sort-Object)
        $adTitles = ($graphUsers | Where-Object { $_.JobTitle } | Select-Object -ExpandProperty JobTitle -Unique | Sort-Object)
        Write-Log "365: $($graphUsers.Count) gebruikers, $($adDepartments.Count) afdelingen, $($adTitles.Count) functies opgehaald"
    } catch {
        Write-Log "Kon gebruikers niet ophalen: $($_.Exception.Message)" "WARN"
    }
}

# Bouw licentie keuze lijst
$licenseChoices = @()
foreach ($lic in ($global:TenantLicenses | Sort-Object SkuPartNumber)) {
    $available = [int]$lic.ActiveUnits - [int]$lic.ConsumedUnits
    $licenseChoices += "$($lic.SkuPartNumber) ($available van $($lic.ActiveUnits) beschikbaar)"
}

# ============================================================================
# STAP 5: Hoofdformulier
# ============================================================================
$main_form = New-Object System.Windows.Forms.Form
$main_form.Text = "Nieuwe gebruiker - $($global:CompanyName) [$($global:EnvironmentMode)]"
$main_form.AutoScaleMode = 'Font'
$main_form.Width = 700; $main_form.Height = 530; $main_form.AutoSize = $true
$main_form.TopMost = $true; $main_form.MinimizeBox = $false; $main_form.MaximizeBox = $false
$main_form.BackColor = $BackColor; $main_form.ForeColor = $ForeColor
$main_form.FormBorderStyle = 'FixedSingle'; $main_form.StartPosition = "CenterScreen"
$main_form.ShowInTaskbar = $true; $main_form.Font = $Form_Font
$main_form.icon = [Drawing.Icon]::ExtractAssociatedIcon((Get-Command powershell).Path)
[void]$main_form.SuspendLayout()

$aSize = [System.Windows.Forms.TextRenderer]::MeasureText("X", $Form_Font)
[int]$gW = $aSize.width + 2; [int]$gH = $aSize.height + 6

# Row 1: Email
$l = New-Object System.Windows.Forms.Label; $l.Text = "Email"; $l.AutoSize = $true; $l.Font = $Form_Font
$l.Location = New-Object System.Drawing.Point((GetXLoc $gW 1), (GetYLoc $gH 1)); [void]$main_form.Controls.Add($l)

$TextBox_User = New-Object System.Windows.Forms.TextBox
$TextBox_User.Location = New-Object System.Drawing.Point((GetXLoc $gW 6 -1), (GetYLoc $gH 1))
$TextBox_User.Width = 200; $TextBox_User.ForeColor = $ForeColor; $TextBox_User.BackColor = $TextBoxBackColor
$TextBox_User.Font = $TextBoxFont; $TextBox_User.TabIndex = 1; $TextBox_User.MaxLength = 64
[void]$main_form.Controls.Add($TextBox_User)

$l = New-Object System.Windows.Forms.Label; $l.Text = "@"; $l.AutoSize = $true; $l.Font = $Form_Font
$l.Location = New-Object System.Drawing.Point((GetXLoc $gW 17 6), (GetYLoc $gH 1)); [void]$main_form.Controls.Add($l)

$ComboBox_DomainName = New-Object System.Windows.Forms.ComboBox
$ComboBox_DomainName.Location = New-Object System.Drawing.Point((GetXLoc $gW 18 6), (GetYLoc $gH 1))
$ComboBox_DomainName.Width = 170; $ComboBox_DomainName.BackColor = $TextBoxBackColor
$ComboBox_DomainName.ForeColor = $ForeColor; $ComboBox_DomainName.FlatStyle = "Flat"
$ComboBox_DomainName.Font = $TextBoxFont; $ComboBox_DomainName.TabIndex = 2
$ComboBox_DomainName.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
foreach ($d in $global:TenantDomains) { [void]$ComboBox_DomainName.Items.Add($d) }
if ($global:PrimaryDomain) { $ComboBox_DomainName.Text = $global:PrimaryDomain }
elseif ($ComboBox_DomainName.Items.Count -gt 0) { $ComboBox_DomainName.SelectedIndex = 0 }
[void]$main_form.Controls.Add($ComboBox_DomainName)

$LabelTick = New-Object System.Windows.Forms.Label; $LabelTick.Text = ""; $LabelTick.AutoSize = $true
$LabelTick.Location = New-Object System.Drawing.Point((GetXLoc $gW 32), (GetYLoc $gH 1 4))
$LabelTick.ForeColor = [System.Drawing.Color]::Green; $LabelTick.Font = $Tick_Font
[void]$main_form.Controls.Add($LabelTick)

# Row 2: Password
$l = New-Object System.Windows.Forms.Label; $l.Text = "Wachtwoord"; $l.AutoSize = $true; $l.Font = $Form_Font
$l.Location = New-Object System.Drawing.Point((GetXLoc $gW 1), (GetYLoc $gH 2)); [void]$main_form.Controls.Add($l)

$TextBox_Password = New-Object System.Windows.Forms.TextBox
$TextBox_Password.Location = New-Object System.Drawing.Point((GetXLoc $gW 6 -1), (GetYLoc $gH 2))
$TextBox_Password.Width = 155; $TextBox_Password.MaxLength = 64; $TextBox_Password.BackColor = $TextBoxBackColor
$TextBox_Password.ForeColor = $ForeColor; $TextBox_Password.Font = $TextBoxFont; $TextBox_Password.TabIndex = 3
$TextBox_Password.Text = CreatePassword 12
$TextBox_Password.Add_DoubleClick({
    $clip = "$($TextBox_User.Text)@$($ComboBox_DomainName.Text)   $($TextBox_Password.Text)"
    if (![string]::IsNullOrWhiteSpace($clip)) { [System.Windows.Forms.Clipboard]::SetText($clip) }
})
[void]$main_form.Controls.Add($TextBox_Password)

$ComboBox_PasswordLength = New-Object System.Windows.Forms.ComboBox
$ComboBox_PasswordLength.Location = New-Object System.Drawing.Point((GetXLoc $gW 14 -6), (GetYLoc $gH 2))
$ComboBox_PasswordLength.Width = 40; $ComboBox_PasswordLength.BackColor = $TextBoxBackColor
$ComboBox_PasswordLength.ForeColor = $ForeColor; $ComboBox_PasswordLength.FlatStyle = "Flat"
$ComboBox_PasswordLength.Font = $TextBoxFont; $ComboBox_PasswordLength.TabIndex = 4
foreach ($n in @("8","10","12","14","16","18","20")) { [void]$ComboBox_PasswordLength.Items.Add($n) }
$ComboBox_PasswordLength.Text = "12"
$ComboBox_PasswordLength.Add_SelectedIndexChanged({ $TextBox_Password.Text = CreatePassword $this.SelectedItem })
[void]$main_form.Controls.Add($ComboBox_PasswordLength)

$l = New-Object System.Windows.Forms.Label; $l.Text = "(tekens)"; $l.AutoSize = $true; $l.Font = $Form_Font
$l.Location = New-Object System.Drawing.Point((GetXLoc $gW 17 6), (GetYLoc $gH 2 -3)); [void]$main_form.Controls.Add($l)

# Row 3: First/Last Name
$l = New-Object System.Windows.Forms.Label; $l.Text = "Voornaam"; $l.AutoSize = $true; $l.Font = $Form_Font
$l.Location = New-Object System.Drawing.Point((GetXLoc $gW 1), (GetYLoc $gH 3)); [void]$main_form.Controls.Add($l)

$TextBox_FirstName = New-Object System.Windows.Forms.TextBox
$TextBox_FirstName.Location = New-Object System.Drawing.Point((GetXLoc $gW 6 -1), (GetYLoc $gH 3))
$TextBox_FirstName.Width = 150; $TextBox_FirstName.BackColor = $TextBoxBackColor; $TextBox_FirstName.ForeColor = $ForeColor
$TextBox_FirstName.Font = $TextBoxFont; $TextBox_FirstName.MaxLength = 64; $TextBox_FirstName.TabIndex = 5
[void]$main_form.Controls.Add($TextBox_FirstName)

$l = New-Object System.Windows.Forms.Label; $l.Text = "Achternaam"; $l.AutoSize = $true; $l.Font = $Form_Font
$l.Location = New-Object System.Drawing.Point((GetXLoc $gW 19), (GetYLoc $gH 3)); [void]$main_form.Controls.Add($l)

$TextBox_LastName = New-Object System.Windows.Forms.TextBox
$TextBox_LastName.Location = New-Object System.Drawing.Point((GetXLoc $gW 24), (GetYLoc $gH 3))
$TextBox_LastName.Width = 150; $TextBox_LastName.BackColor = $TextBoxBackColor; $TextBox_LastName.ForeColor = $ForeColor
$TextBox_LastName.Font = $TextBoxFont; $TextBox_LastName.MaxLength = 64; $TextBox_LastName.TabIndex = 6
[void]$main_form.Controls.Add($TextBox_LastName)

# Row 4: Display Name + Setup Like
$l = New-Object System.Windows.Forms.Label; $l.Text = "Weergavenaam"; $l.AutoSize = $true; $l.Font = $Form_Font
$l.Location = New-Object System.Drawing.Point((GetXLoc $gW 1), (GetYLoc $gH 4)); [void]$main_form.Controls.Add($l)

$TextBox_DisplayName = New-Object System.Windows.Forms.TextBox
$TextBox_DisplayName.Location = New-Object System.Drawing.Point((GetXLoc $gW 6 -1), (GetYLoc $gH 4))
$TextBox_DisplayName.Width = 200; $TextBox_DisplayName.BackColor = $TextBoxBackColor; $TextBox_DisplayName.ForeColor = $ForeColor
$TextBox_DisplayName.Font = $TextBoxFont; $TextBox_DisplayName.MaxLength = 256; $TextBox_DisplayName.TabIndex = 7
[void]$main_form.Controls.Add($TextBox_DisplayName)

$l = New-Object System.Windows.Forms.Label; $l.Text = "Setup Like"; $l.AutoSize = $true; $l.Font = $Form_Font
$l.Location = New-Object System.Drawing.Point((GetXLoc $gW 19), (GetYLoc $gH 4)); [void]$main_form.Controls.Add($l)

$ComboBox_BasedOn = New-Object System.Windows.Forms.ComboBox
$ComboBox_BasedOn.Location = New-Object System.Drawing.Point((GetXLoc $gW 24), (GetYLoc $gH 4))
$ComboBox_BasedOn.Width = 200; $ComboBox_BasedOn.BackColor = $TextBoxBackColor; $ComboBox_BasedOn.ForeColor = $ForeColor
$ComboBox_BasedOn.FlatStyle = "Flat"; $ComboBox_BasedOn.Font = $TextBoxFont; $ComboBox_BasedOn.TabIndex = 8
$ComboBox_BasedOn.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$ComboBox_BasedOn.Items.Add("")
foreach ($n in $adNames) { [void]$ComboBox_BasedOn.Items.Add($n) }
[void]$main_form.Controls.Add($ComboBox_BasedOn)

# Row 5: Department + Manager
$l = New-Object System.Windows.Forms.Label; $l.Text = "Afdeling"; $l.AutoSize = $true; $l.Font = $Form_Font
$l.Location = New-Object System.Drawing.Point((GetXLoc $gW 1), (GetYLoc $gH 5)); [void]$main_form.Controls.Add($l)

$ComboBox_Department = New-Object System.Windows.Forms.ComboBox
$ComboBox_Department.Location = New-Object System.Drawing.Point((GetXLoc $gW 6 -1), (GetYLoc $gH 5))
$ComboBox_Department.Width = 200; $ComboBox_Department.BackColor = $TextBoxBackColor; $ComboBox_Department.ForeColor = $ForeColor
$ComboBox_Department.FlatStyle = "Flat"; $ComboBox_Department.Font = $TextBoxFont; $ComboBox_Department.TabIndex = 9
foreach ($d in $adDepartments) { [void]$ComboBox_Department.Items.Add($d) }
[void]$main_form.Controls.Add($ComboBox_Department)

$l = New-Object System.Windows.Forms.Label; $l.Text = "Manager"; $l.AutoSize = $true; $l.Font = $Form_Font
$l.Location = New-Object System.Drawing.Point((GetXLoc $gW 19), (GetYLoc $gH 5)); [void]$main_form.Controls.Add($l)

$ComboBox_Manager = New-Object System.Windows.Forms.ComboBox
$ComboBox_Manager.Location = New-Object System.Drawing.Point((GetXLoc $gW 24), (GetYLoc $gH 5))
$ComboBox_Manager.Width = 200; $ComboBox_Manager.BackColor = $TextBoxBackColor; $ComboBox_Manager.ForeColor = $ForeColor
$ComboBox_Manager.FlatStyle = "Flat"; $ComboBox_Manager.Font = $TextBoxFont; $ComboBox_Manager.TabIndex = 10
$ComboBox_Manager.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$ComboBox_Manager.Items.Add("")
foreach ($n in $adNames) { [void]$ComboBox_Manager.Items.Add($n) }
[void]$main_form.Controls.Add($ComboBox_Manager)

# Row 6: Title
$l = New-Object System.Windows.Forms.Label; $l.Text = "Functie"; $l.AutoSize = $true; $l.Font = $Form_Font
$l.Location = New-Object System.Drawing.Point((GetXLoc $gW 1), (GetYLoc $gH 6)); [void]$main_form.Controls.Add($l)

$ComboBox_Title = New-Object System.Windows.Forms.ComboBox
$ComboBox_Title.Location = New-Object System.Drawing.Point((GetXLoc $gW 6 -1), (GetYLoc $gH 6))
$ComboBox_Title.Width = 300; $ComboBox_Title.BackColor = $TextBoxBackColor; $ComboBox_Title.ForeColor = $ForeColor
$ComboBox_Title.FlatStyle = "Flat"; $ComboBox_Title.Font = $TextBoxFont; $ComboBox_Title.TabIndex = 11
foreach ($t in $adTitles) { [void]$ComboBox_Title.Items.Add($t) }
[void]$main_form.Controls.Add($ComboBox_Title)

# Row 7: Locatie (UsageLocation)
$l = New-Object System.Windows.Forms.Label; $l.Text = "Locatie"; $l.AutoSize = $true; $l.Font = $Form_Font
$l.Location = New-Object System.Drawing.Point((GetXLoc $gW 1), (GetYLoc $gH 7)); [void]$main_form.Controls.Add($l)

$ComboBox_Location = New-Object System.Windows.Forms.ComboBox
$ComboBox_Location.Location = New-Object System.Drawing.Point((GetXLoc $gW 6 -1), (GetYLoc $gH 7))
$ComboBox_Location.Width = 200; $ComboBox_Location.BackColor = $TextBoxBackColor; $ComboBox_Location.ForeColor = $ForeColor
$ComboBox_Location.FlatStyle = "Flat"; $ComboBox_Location.Font = $TextBoxFont; $ComboBox_Location.TabIndex = 12
$ComboBox_Location.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$countries = @(
    "NL - Nederland", "BE - Belgie", "DE - Duitsland", "FR - Frankrijk", "GB - Verenigd Koninkrijk",
    "US - Verenigde Staten", "CA - Canada", "AU - Australie", "NZ - Nieuw-Zeeland",
    "AT - Oostenrijk", "CH - Zwitserland", "DK - Denemarken", "ES - Spanje", "FI - Finland",
    "IE - Ierland", "IT - Italie", "LU - Luxemburg", "NO - Noorwegen", "PL - Polen",
    "PT - Portugal", "SE - Zweden", "ZA - Zuid-Afrika", "IN - India", "JP - Japan",
    "SG - Singapore", "PH - Filipijnen", "BR - Brazilie", "MX - Mexico"
)
foreach ($c in $countries) { [void]$ComboBox_Location.Items.Add($c) }
$ComboBox_Location.SelectedIndex = 0
Set-Tooltip $ComboBox_Location "UsageLocation - verplicht voor het toewijzen van licenties"
[void]$main_form.Controls.Add($ComboBox_Location)

# Row 8: Extra SMTP adressen (alleen On-Prem)
$TextBox_ExtraSMTP = $null
if ($global:EnvironmentMode -eq "OnPrem") {
    $l = New-Object System.Windows.Forms.Label; $l.Text = "Extra SMTP"; $l.AutoSize = $true; $l.Font = $Form_Font
    $l.Location = New-Object System.Drawing.Point((GetXLoc $gW 1), (GetYLoc $gH 8)); [void]$main_form.Controls.Add($l)

    $TextBox_ExtraSMTP = New-Object System.Windows.Forms.TextBox
    $TextBox_ExtraSMTP.Location = New-Object System.Drawing.Point((GetXLoc $gW 6 -1), (GetYLoc $gH 8))
    $TextBox_ExtraSMTP.Width = 400; $TextBox_ExtraSMTP.BackColor = $TextBoxBackColor
    $TextBox_ExtraSMTP.ForeColor = [System.Drawing.Color]::Gray; $TextBox_ExtraSMTP.Font = $TextBoxFont
    $TextBox_ExtraSMTP.Text = "extra@domein.nl; alias@domein.nl"
    $TextBox_ExtraSMTP.Add_GotFocus({
        if ($TextBox_ExtraSMTP.Text -eq "extra@domein.nl; alias@domein.nl") {
            $TextBox_ExtraSMTP.Text = ""; $TextBox_ExtraSMTP.ForeColor = $ForeColor
        }
    })
    $TextBox_ExtraSMTP.Add_LostFocus({
        if ([string]::IsNullOrWhiteSpace($TextBox_ExtraSMTP.Text)) {
            $TextBox_ExtraSMTP.Text = "extra@domein.nl; alias@domein.nl"; $TextBox_ExtraSMTP.ForeColor = [System.Drawing.Color]::Gray
        }
    })
    Set-Tooltip $TextBox_ExtraSMTP "Extra SMTP proxy adressen, gescheiden door ; (worden als smtp: alias toegevoegd)"
    [void]$main_form.Controls.Add($TextBox_ExtraSMTP)
}

# Row 9: AD Groepen (On-Prem) of 365 Licentie (Cloud Only)
if ($global:EnvironmentMode -eq "OnPrem") {
    $l = New-Object System.Windows.Forms.Label; $l.Text = "AD Groepen"; $l.AutoSize = $true; $l.Font = $Form_Font
    $l.Location = New-Object System.Drawing.Point((GetXLoc $gW 1), (GetYLoc $gH 9)); [void]$main_form.Controls.Add($l)

    # Haal alle AD groepen op (exclusief systeem/builtin groepen)
    $global:AllADGroups = @()
    $excludeGroups = @(
        'Domain Admins', 'Domain Controllers', 'Domain Computers', 'Domain Guests',
        'Domain Users', 'Enterprise Admins', 'Enterprise Read-only Domain Controllers',
        'Schema Admins', 'Group Policy Creator Owners', 'Cloneable Domain Controllers',
        'Protected Users', 'Key Admins', 'Enterprise Key Admins', 'DnsAdmins',
        'DnsUpdateProxy', 'Read-only Domain Controllers', 'Cert Publishers',
        'RAS and IAS Servers', 'Allowed RODC Password Replication Group',
        'Denied RODC Password Replication Group', 'DHCP Administrators', 'DHCP Users',
        'WinRMRemoteWMIUsers__', 'Access Control Assistance Operators',
        'Account Operators', 'Administrators', 'Backup Operators',
        'Certificate Service DCOM Access', 'Cryptographic Operators',
        'Distributed COM Users', 'Event Log Readers', 'Guests',
        'Hyper-V Administrators', 'IIS_IUSRS', 'Incoming Forest Trust Builders',
        'Network Configuration Operators', 'Performance Log Users',
        'Performance Monitor Users', 'Pre-Windows 2000 Compatible Access',
        'Print Operators', 'Remote Desktop Users', 'Remote Management Users',
        'Replicator', 'Server Operators', 'Storage Replica Administrators',
        'Terminal Server License Servers', 'Users', 'Windows Authorization Access Group'
    )
    try {
        $global:AllADGroups = Get-ADGroup -Filter * -Properties GroupScope, GroupCategory -ErrorAction SilentlyContinue |
            Where-Object {
                $_.Name -notin $excludeGroups -and
                $_.DistinguishedName -notlike "*CN=Builtin,*" -and
                $_.Name -notlike "SMS*" -and
                $_.Name -notlike "DVDR*" -and
                $_.Name -notlike "SQLServer*"
            } | Sort-Object Name
    } catch { Write-Log "AD groepen ophalen mislukt: $($_.Exception.Message)" "WARN" }

    # Zoekbalk
    $TextBox_GroupSearch = New-Object System.Windows.Forms.TextBox
    $TextBox_GroupSearch.Location = New-Object System.Drawing.Point((GetXLoc $gW 6 -1), (GetYLoc $gH 9))
    $TextBox_GroupSearch.Width = 400; $TextBox_GroupSearch.Height = 22
    $TextBox_GroupSearch.BackColor = $TextBoxBackColor; $TextBox_GroupSearch.ForeColor = [System.Drawing.Color]::Gray
    $TextBox_GroupSearch.Font = $TextBoxFont; $TextBox_GroupSearch.Text = "Zoek groepen..."
    $TextBox_GroupSearch.Add_GotFocus({
        if ($TextBox_GroupSearch.Text -eq "Zoek groepen...") {
            $TextBox_GroupSearch.Text = ""; $TextBox_GroupSearch.ForeColor = $ForeColor
        }
    })
    $TextBox_GroupSearch.Add_LostFocus({
        if ([string]::IsNullOrWhiteSpace($TextBox_GroupSearch.Text)) {
            $TextBox_GroupSearch.Text = "Zoek groepen..."; $TextBox_GroupSearch.ForeColor = [System.Drawing.Color]::Gray
        }
    })
    [void]$main_form.Controls.Add($TextBox_GroupSearch)

    # Groepen listbox
    $ListBox_ADGroups = New-Object System.Windows.Forms.CheckedListBox
    $ListBox_ADGroups.Location = New-Object System.Drawing.Point((GetXLoc $gW 6 -1), ((GetYLoc $gH 9) + 26))
    $ListBox_ADGroups.Width = 400; $ListBox_ADGroups.Height = 100
    $ListBox_ADGroups.BackColor = $TextBoxBackColor; $ListBox_ADGroups.ForeColor = $ForeColor
    $ListBox_ADGroups.Font = $TextBoxFont; $ListBox_ADGroups.TabIndex = 13
    $ListBox_ADGroups.CheckOnClick = $true
    $ListBox_ADGroups.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    foreach ($g in $global:AllADGroups) { [void]$ListBox_ADGroups.Items.Add($g.Name) }
    Set-Tooltip $ListBox_ADGroups "Vink groepen aan om de gebruiker aan toe te voegen (naast Setup Like kopie)"
    [void]$main_form.Controls.Add($ListBox_ADGroups)

    # Bewaar aangevinkte groepen bij filteren
    $global:CheckedGroupNames = [System.Collections.Generic.HashSet[string]]::new()

    $TextBox_GroupSearch.Add_TextChanged({
        $searchText = $TextBox_GroupSearch.Text
        if ($searchText -eq "Zoek groepen...") { return }

        # Bewaar huidige selecties
        for ($i = 0; $i -lt $ListBox_ADGroups.Items.Count; $i++) {
            $itemName = $ListBox_ADGroups.Items[$i].ToString()
            if ($ListBox_ADGroups.GetItemChecked($i)) { [void]$global:CheckedGroupNames.Add($itemName) }
            else { [void]$global:CheckedGroupNames.Remove($itemName) }
        }

        # Filter en herbouw lijst
        $ListBox_ADGroups.Items.Clear()
        foreach ($g in $global:AllADGroups) {
            if ([string]::IsNullOrWhiteSpace($searchText) -or $g.Name -like "*$searchText*") {
                [void]$ListBox_ADGroups.Items.Add($g.Name)
                $idx = $ListBox_ADGroups.Items.Count - 1
                if ($global:CheckedGroupNames.Contains($g.Name)) {
                    $ListBox_ADGroups.SetItemChecked($idx, $true)
                }
            }
        }
    })

    # Label met telling
    $lblGroupCount = New-Object System.Windows.Forms.Label
    $lblGroupCount.Text = "$($global:AllADGroups.Count) groepen"
    $lblGroupCount.AutoSize = $true; $lblGroupCount.Font = $Font8
    $lblGroupCount.ForeColor = [System.Drawing.Color]::Gray
    $lblGroupCount.Location = New-Object System.Drawing.Point(((GetXLoc $gW 6 -1) + 310), (GetYLoc $gH 9))
    [void]$main_form.Controls.Add($lblGroupCount)

    # Verberg licentie in OnPrem - wordt later geselecteerd na sync
    $ComboBox_License = $null
    $main_form.Height = 660
} else {
    $l = New-Object System.Windows.Forms.Label; $l.Text = "365 Licentie"; $l.AutoSize = $true; $l.Font = $Form_Font
    $l.Location = New-Object System.Drawing.Point((GetXLoc $gW 1), (GetYLoc $gH 8)); [void]$main_form.Controls.Add($l)

    $ComboBox_License = New-Object System.Windows.Forms.CheckedListBox
    $ComboBox_License.Location = New-Object System.Drawing.Point((GetXLoc $gW 6 -1), (GetYLoc $gH 8))
    $ComboBox_License.Width = 400; $ComboBox_License.Height = 80
    $ComboBox_License.BackColor = $TextBoxBackColor; $ComboBox_License.ForeColor = $ForeColor
    $ComboBox_License.Font = $TextBoxFont; $ComboBox_License.TabIndex = 13
    $ComboBox_License.CheckOnClick = $true; $ComboBox_License.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    foreach ($lc in $licenseChoices) { [void]$ComboBox_License.Items.Add($lc) }
    [void]$main_form.Controls.Add($ComboBox_License)
    $ListBox_ADGroups = $null
}

# Auto-fill logica
$TextBox_User.Add_TextChanged({
    $sel = $TextBox_User.SelectionStart
    $TextBox_User.Text = $TextBox_User.Text.ToLower().TrimEnd()
    $split = $TextBox_User.Text -split "\."
    if ($split.length -eq 2) {
        if (-not $split[0].Contains("@")) { $TextBox_FirstName.Text = (Get-Culture).TextInfo.ToTitleCase($split[0]) }
        if (-not $split[1].Contains("@")) { $TextBox_LastName.Text = (Get-Culture).TextInfo.ToTitleCase($split[1]) }
        $TextBox_DisplayName.Text = "$($TextBox_FirstName.Text) $($TextBox_LastName.Text)"
        if ($split[1].length -gt 0) { $LabelTick.Text = [Char]8730; $LabelTick.ForeColor = [System.Drawing.Color]::Green }
        else { $LabelTick.Text = "" }
        if ($global:ADusers) {
            if ($global:EnvironmentMode -eq "OnPrem") { $ex = $global:ADusers | Where-Object { $_.SamAccountName -like $TextBox_User.Text } }
            else { $ex = $global:ADusers | Where-Object { $_.UserPrincipalName -like "$($TextBox_User.Text)@*" } }
            if ($ex) { $LabelTick.Text = "X"; $LabelTick.ForeColor = [System.Drawing.Color]::Red }
        }
    } else {
        if (-not $split[0].Contains("@")) { $TextBox_FirstName.Text = (Get-Culture).TextInfo.ToTitleCase($split[0]) }
        $LabelTick.Text = ""
    }
    if ($sel -lt $TextBox_User.Text.Length) { $TextBox_User.Select($sel, 0) }
    else { $TextBox_User.Select($TextBox_User.Text.Length, 0) }
})

# Buttons
[int]$btnY = [int]$main_form.Height - 72

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Size(390, $btnY); $okButton.Size = New-Object System.Drawing.Size(130, 28)
$okButton.Text = "Aanmaken"; $okButton.FlatStyle = "Flat"
$okButton.FlatAppearance.BorderColor = $ButtonBorderColor; $okButton.FlatAppearance.BorderSize = 2
$okButton.FlatAppearance.MouseDownBackColor = $ButtonMouseDownColor; $okButton.FlatAppearance.MouseOverBackColor = $ButtonMouseOverColor
$okButton.ForeColor = $ForeColor; $okButton.BackColor = $ButtonColor; $okButton.Font = $ButtonFont
$okButton.Add_Click({
    if ([string]::IsNullOrWhiteSpace($TextBox_User.Text)) { [System.Windows.Forms.MessageBox]::Show("Vul een gebruikersnaam in.", "Fout", 0, 16); return }
    if ([string]::IsNullOrWhiteSpace($TextBox_FirstName.Text) -or [string]::IsNullOrWhiteSpace($TextBox_LastName.Text)) { [System.Windows.Forms.MessageBox]::Show("Vul voor- en achternaam in.", "Fout", 0, 16); return }
    if ([string]::IsNullOrWhiteSpace($TextBox_Password.Text)) { [System.Windows.Forms.MessageBox]::Show("Vul een wachtwoord in.", "Fout", 0, 16); return }
    $main_form.Tag = "OK"; $main_form.Close()
})
[void]$main_form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Size(140, $btnY); $cancelButton.Size = New-Object System.Drawing.Size(130, 28)
$cancelButton.Text = "Annuleren"; $cancelButton.FlatStyle = "Flat"
$cancelButton.FlatAppearance.BorderColor = $ButtonBorderColor; $cancelButton.FlatAppearance.BorderSize = 2
$cancelButton.FlatAppearance.MouseDownBackColor = $ButtonMouseDownColor; $cancelButton.FlatAppearance.MouseOverBackColor = $ButtonMouseOverColor
$cancelButton.ForeColor = $ForeColor; $cancelButton.BackColor = $ButtonColor; $cancelButton.Font = $ButtonFont
$cancelButton.Add_Click({ $main_form.Tag = $null; $main_form.Close() })
[void]$main_form.Controls.Add($cancelButton)

[void]$main_form.ResumeLayout()
[void]$main_form.ShowDialog()

if ($main_form.Tag -ne "OK") { try { Disconnect-MgGraph -ErrorAction Ignore > $null } catch {}; exit }

# ============================================================================
# STAP 6: Gebruiker aanmaken
# ============================================================================
$loginname = $TextBox_User.Text.Trim()
$firstname = $TextBox_FirstName.Text.Trim()
$lastname = $TextBox_LastName.Text.Trim()
$displayname = $TextBox_DisplayName.Text.Trim()
$password = $TextBox_Password.Text.Trim()
$emailaddress = "$($loginname)@$($ComboBox_DomainName.Text)".ToLower()
$department = $ComboBox_Department.Text
$title = $ComboBox_Title.Text
$manager = ""; if ($ComboBox_Manager.SelectedItem) { $manager = $ComboBox_Manager.SelectedItem }
$basedOn = ""; if ($ComboBox_BasedOn.SelectedItem) { $basedOn = $ComboBox_BasedOn.SelectedItem }
$usageLocation = ($ComboBox_Location.SelectedItem.ToString() -split " - ")[0].Trim()
# Handmatig geselecteerde AD groepen ophalen (inclusief weggefilterde items)
$manualADGroups = @()
if ($ListBox_ADGroups) {
    # Eerst: bewaar huidige zichtbare selecties in de HashSet
    for ($i = 0; $i -lt $ListBox_ADGroups.Items.Count; $i++) {
        $itemName = $ListBox_ADGroups.Items[$i].ToString()
        if ($ListBox_ADGroups.GetItemChecked($i)) { [void]$global:CheckedGroupNames.Add($itemName) }
        else { [void]$global:CheckedGroupNames.Remove($itemName) }
    }
    # Dan: alle aangevinkte groepen (ook weggefilterde)
    $manualADGroups = @($global:CheckedGroupNames)
}
$SecurePassword = ConvertTo-SecureString $password -AsPlainText -Force
$errorOccured = $false; $errorMessage = ""; $created = $false; $NewGroups = ""
# Extra SMTP proxy adressen
$extraSMTPAddresses = @()
if ($TextBox_ExtraSMTP -and $TextBox_ExtraSMTP.Text -and $TextBox_ExtraSMTP.Text -ne "extra@domein.nl; alias@domein.nl") {
    $extraSMTPAddresses = $TextBox_ExtraSMTP.Text -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '@' }
    if ($extraSMTPAddresses.Count -gt 0) {
        Write-Log "Extra SMTP adressen: $($extraSMTPAddresses -join ', ')"
    }
}

Write-Host "`n======================================== " -ForegroundColor Cyan
Write-Host "Gebruiker aanmaken: $emailaddress" -ForegroundColor Cyan
Write-Host "Modus: $($global:EnvironmentMode)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# --- CLOUD ONLY ---
if ($global:EnvironmentMode -eq "CloudOnly") {
    Write-Log "=== START Cloud Only user aanmaken: $emailaddress ==="
    Write-Log "Gegevens: Naam=$displayname | Afdeling=$department | Functie=$title | Manager=$manager | BasedOn=$basedOn | Locatie=$usageLocation"
    $existing = $null
    try { $existing = Get-MgUser -Filter "UserPrincipalName eq '$emailaddress'" -ErrorAction SilentlyContinue } catch {}

    if ($existing) {
        $errorMessage = "Gebruiker $emailaddress bestaat al in Office 365!"
        $errorOccured = $true
        Write-Log $errorMessage "ERROR"
    } else {
        Write-Log "Gebruiker aanmaken in Office 365..."
        $pwProfile = @{ Password = $password; ForceChangePasswordNextSignIn = $false }
        $ht = @{
            UserPrincipalName = $emailaddress; DisplayName = $displayname; PasswordProfile = $pwProfile
            MailNickName = $loginname; GivenName = $firstname; Surname = $lastname
            JobTitle = $title; Department = $department; AccountEnabled = $true
            UsageLocation = $usageLocation; ErrorAction = 'Stop'
        }
        $keysToRemove = $ht.keys | Where-Object { $null -eq $ht[$_] -or ($ht[$_] -is [string] -and [string]::IsNullOrWhiteSpace($ht[$_])) }
        foreach ($k in @($keysToRemove)) { $ht.Remove($k) }

        try {
            $newUser = New-MgUser @ht; $created = $true
            Write-Log "Gebruiker $emailaddress aangemaakt in Office 365 (Id: $($newUser.Id))" "SUCCESS"
            Start-Sleep -Seconds 2
        } catch {
            $errorMessage = "Fout bij aanmaken: $($_.Exception.Message)"; $errorOccured = $true
            Write-Log $errorMessage "ERROR"
        }

        # Manager instellen
        if ($created -and $manager) {
            Write-Log "Manager instellen: $manager"
            try {
                $mgrUser = $global:ADusers | Where-Object { $_.DisplayName -eq $manager }
                if ($mgrUser) {
                    Set-MgUserManagerByRef -UserId $emailaddress -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($mgrUser.Id)" }
                    Write-Log "Manager ingesteld: $manager" "SUCCESS"
                } else {
                    Write-Log "Manager '$manager' niet gevonden in gebruikerslijst" "WARN"
                }
            } catch { Write-Log "Manager instellen mislukt: $($_.Exception.Message)" "ERROR" }
        }

        # Licenties toewijzen (meerdere mogelijk)
        $licenseAssigned = $false
        if ($created -and $ComboBox_License -and $ComboBox_License.CheckedItems.Count -gt 0) {
            foreach ($selLic in $ComboBox_License.CheckedItems) {
                $skuPart = ($selLic.ToString() -split " \(")[0].Trim()
                $sku = $global:TenantLicenses | Where-Object { $_.SkuPartNumber -eq $skuPart }
                if ($sku) {
                    Write-Log "Licentie toewijzen: $skuPart (SkuId: $($sku.SkuId))"
                    try {
                        Set-MgUserLicense -UserId $emailaddress -BodyParameter @{ AddLicenses = @(@{ DisabledPlans = @(); SkuId = $sku.SkuId }); RemoveLicenses = @() }
                        Write-Log "Licentie '$skuPart' toegewezen" "SUCCESS"
                        $licenseAssigned = $true
                    } catch { Write-Log "Licentie toewijzen mislukt: $($_.Exception.Message)" "ERROR" }
                } else {
                    Write-Log "Licentie SKU '$skuPart' niet gevonden in tenant" "ERROR"
                }
            }
        } elseif ($created) {
            Write-Log "Geen licentie geselecteerd - overgeslagen"
        }

        # Groepen kopieren van BasedOn gebruiker
        if ($created -and $basedOn) {
            $copyGroups = [System.Windows.Forms.MessageBox]::Show(
                "Wil je de 365 groepen van '$basedOn' kopieren naar $displayname ?",
                "365 Groepen kopieren",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($copyGroups -eq [System.Windows.Forms.DialogResult]::Yes) {
            Write-Log "=== Groepen kopieren van '$basedOn' ==="
            $groupsOK = 0; $groupsFailed = 0; $groupsSkipped = 0
            try {
                $srcUser = $global:ADusers | Where-Object { $_.DisplayName -eq $basedOn }
                if ($srcUser) {
                    Write-Log "Brongebruiker gevonden: $($srcUser.DisplayName) (Id: $($srcUser.Id))"
                    $grps = Get-MgUserMemberOf -UserId $srcUser.Id -All
                    Write-Log "Brongebruiker is lid van $($grps.Count) groepen/rollen"
                    $tgtUser = Get-MgUser -Filter "UserPrincipalName eq '$emailaddress'"
                    foreach ($g in $grps) {
                        # Alleen groepen verwerken, geen directory roles
                        if ($g.AdditionalProperties.'@odata.type' -ne '#microsoft.graph.group') {
                            $typeName = $g.AdditionalProperties.'@odata.type'
                            Write-Log "  [SKIP] '$($g.Id)' is geen groep (type: $typeName)"
                            $groupsSkipped++
                            continue
                        }
                        $groupName = $g.AdditionalProperties.displayName
                        $groupTypes = $g.AdditionalProperties.groupTypes -join ","
                        $secEnabled = $g.AdditionalProperties.securityEnabled
                        $mailEnabled = $g.AdditionalProperties.mailEnabled
                        try {
                            New-MgGroupMember -GroupId $g.Id -DirectoryObjectId $tgtUser.Id -ErrorAction Stop | Out-Null
                            $NewGroups += if ($NewGroups) { ", $groupName" } else { $groupName }
                            Write-Log "  [OK] Toegevoegd aan groep: $groupName (security=$secEnabled, mail=$mailEnabled, types=$groupTypes)" "SUCCESS"
                            $groupsOK++
                        } catch {
                            $errMsg = $_.Exception.Message
                            if ($errMsg -like "*already exist*") {
                                Write-Log "  [SKIP] '$groupName' - gebruiker is al lid" "WARN"
                                $groupsSkipped++
                            } else {
                                Write-Log "  [FAIL] Groep '$groupName' mislukt: $errMsg" "ERROR"
                                $groupsFailed++
                            }
                        }
                    }
                    Write-Log "=== Groepen resultaat: $groupsOK toegevoegd, $groupsFailed mislukt, $groupsSkipped overgeslagen ==="
                } else {
                    Write-Log "Brongebruiker '$basedOn' niet gevonden" "ERROR"
                }
            } catch { Write-Log "Groepen kopieren mislukt: $($_.Exception.Message)" "ERROR" }
            } else {
                Write-Log "365 groepen kopieren overgeslagen door gebruiker"
            }
        }

        # 365 Inrichtingsformulier - handmatige groepen, DLs en shared mailboxes
        if ($created -and $newUser) {
            Write-Log "365 inrichtingsformulier tonen..."
            $selections365 = Show-365ConfigForm -UserDisplayName $displayname -UserEmail $emailaddress -UserId $newUser.Id
            if ($selections365) {
                $cfgResult = Process-365Selections -Selections $selections365 -UserId $newUser.Id -UserEmail $emailaddress -ExoAvailable $selections365.ExoAvailable -LicenseAssigned $licenseAssigned
                if ($cfgResult -and $cfgResult.OK -gt 0) {
                    Write-Log "365 handmatige inrichting: $($cfgResult.OK) items toegevoegd" "SUCCESS"
                }
            }
        }
    }
    Write-Log "=== EINDE Cloud Only verwerking ==="
}

# --- ON-PREM ---
if ($global:EnvironmentMode -eq "OnPrem") {
    Write-Log "=== START On-Prem user aanmaken: $emailaddress ==="
    Write-Log "Gegevens: Naam=$displayname | Afdeling=$department | Functie=$title | Manager=$manager | BasedOn=$basedOn | Locatie=$usageLocation | OU=$OUPath"
    if ($manualADGroups.Count -gt 0) { Write-Log "Handmatige AD groepen: $($manualADGroups -join ', ')" }

    # Admin check
    $IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
    Write-Log "PowerShell draait als Administrator: $IsAdmin"
    if (-not $IsAdmin) {
        Write-Log "LET OP: PowerShell draait NIET als Administrator - dit kan Access Denied veroorzaken!" "WARN"
    }

    $checkAD = $global:ADusers | Where-Object { $_.SamAccountName -like $loginname }

    if ($checkAD) {
        $errorMessage = "Gebruiker $loginname bestaat al in Active Directory!"
        $errorOccured = $true
        Write-Log $errorMessage "ERROR"
    } else {
        Write-Log "Gebruiker aanmaken in Active Directory (lokaal)..."
        # Gebruik de lokale DC als server
        $dcServer = $env:COMPUTERNAME
        Write-Log "Doel DC: $dcServer"

        $adHash = @{
            SamAccountName = $loginname; UserPrincipalName = $emailaddress; DisplayName = $displayname
            Name = $displayname; GivenName = $firstname; Surname = $lastname
            EmailAddress = $emailaddress
            Title = $title; Department = $department; Company = $global:CompanyName
            Enabled = $false; Server = $dcServer
        }
        if ($OUPath) { $adHash['Path'] = $OUPath }
        if ($manager) {
            $mgrUser = $global:ADusers | Where-Object { $_.Name -eq $manager }
            if ($mgrUser) {
                $adHash['Manager'] = $mgrUser.DistinguishedName
                Write-Log "Manager DN: $($mgrUser.DistinguishedName)"
            } else {
                Write-Log "Manager '$manager' niet gevonden in AD" "WARN"
            }
        }
        # Verwijder lege string velden (maar behoud booleans)
        $keysToRemove = $adHash.keys | Where-Object { $null -eq $adHash[$_] -or ($adHash[$_] -is [string] -and [string]::IsNullOrWhiteSpace($adHash[$_])) }
        foreach ($k in @($keysToRemove)) { $adHash.Remove($k) }

        try {
            # Stap 1: User aanmaken (disabled, zonder wachtwoord)
            Write-Log "New-ADUser parameters: $($adHash.Keys -join ', ')"
            New-ADUser @adHash -ErrorAction Stop
            Write-Log "Gebruiker '$loginname' aangemaakt in AD (disabled)" "SUCCESS"

            # Stap 2: Wachtwoord instellen
            try {
                Set-ADAccountPassword -Identity $loginname -NewPassword $SecurePassword -Reset -Server $dcServer -ErrorAction Stop
                Write-Log "Wachtwoord ingesteld" "SUCCESS"
            } catch {
                Write-Log "Wachtwoord instellen mislukt: $($_.Exception.Message)" "ERROR"
                Write-Log "Controleer of het wachtwoord voldoet aan het wachtwoordbeleid (minimaal 8 tekens, hoofdletter, kleine letter, cijfer, speciaal teken)" "WARN"
            }

            # Stap 3: Account activeren
            try {
                Enable-ADAccount -Identity $loginname -Server $dcServer -ErrorAction Stop
                Write-Log "Account geactiveerd" "SUCCESS"
                $created = $true
            } catch {
                Write-Log "Account activeren mislukt: $($_.Exception.Message)" "ERROR"
                Write-Log "Mogelijk voldoet het wachtwoord niet aan het domein wachtwoordbeleid" "WARN"
                $created = $true  # User bestaat, maar is disabled
            }

            # Stap 4: Wachtwoord flags apart instellen
            try {
                Set-ADUser -Identity $loginname -PasswordNeverExpires $true -CannotChangePassword $false -ChangePasswordAtLogon $false -Server $dcServer -ErrorAction Stop
                Write-Log "Wachtwoord flags ingesteld" "SUCCESS"
            } catch {
                Write-Log "Wachtwoord flags instellen mislukt: $($_.Exception.Message)" "WARN"
            }

            # Stap 5: ProxyAddresses apart instellen
            try {
                Set-ADUser -Identity $loginname -Add @{proxyAddresses = "SMTP:$emailaddress"} -Server $dcServer -ErrorAction Stop
                Write-Log "ProxyAddress ingesteld: SMTP:$emailaddress" "SUCCESS"

                # Extra SMTP adressen toevoegen (als alias, lowercase smtp:)
                if ($extraSMTPAddresses.Count -gt 0) {
                    foreach ($extraAddr in $extraSMTPAddresses) {
                        try {
                            Set-ADUser -Identity $loginname -Add @{proxyAddresses = "smtp:$extraAddr"} -Server $dcServer -ErrorAction Stop
                            Write-Log "Extra proxy: smtp:$extraAddr" "SUCCESS"
                        } catch {
                            Write-Log "Extra proxy mislukt voor $extraAddr : $($_.Exception.Message)" "WARN"
                        }
                    }
                }
            } catch {
                Write-Log "ProxyAddress instellen mislukt: $($_.Exception.Message)" "WARN"
            }
        } catch {
            $errDetail = $_.Exception.Message
            if ($_.Exception.InnerException) { $errDetail += " | Inner: $($_.Exception.InnerException.Message)" }
            $errorMessage = "AD aanmaak mislukt: $errDetail"
            $errorOccured = $true
            Write-Log $errorMessage "ERROR"
            if ($errDetail -like "*Access*denied*" -or $errDetail -like "*Access is denied*") {
                Write-Log "Mogelijke oorzaken: onvoldoende rechten op de OU, of wachtwoord voldoet niet aan het domein wachtwoordbeleid" "WARN"
            }
        }

        # AD Groepen kopieren van BasedOn gebruiker
        if ($created -and $basedOn) {
            Write-Log "=== AD Groepen kopieren van '$basedOn' ==="
            $groupsOK = 0; $groupsFailed = 0
            try {
                $srcUser = $global:ADusers | Where-Object { $_.Name -eq $basedOn }
                if ($srcUser) {
                    Write-Log "Brongebruiker gevonden: $($srcUser.Name) (SAM: $($srcUser.SamAccountName))"
                    $srcGroups = Get-ADPrincipalGroupMembership -Identity $srcUser.SamAccountName -Server $dcServer -ErrorAction Stop |
                        Where-Object { $_.Name -ne 'Domain Users' }
                    Write-Log "Brongebruiker is lid van $($srcGroups.Count) groepen (excl. Domain Users)"
                    foreach ($grp in $srcGroups) {
                        try {
                            Add-ADGroupMember -Identity $grp.Name -Members $loginname -Server $dcServer -ErrorAction Stop
                            $NewGroups += if ($NewGroups) { ", $($grp.Name)" } else { $grp.Name }
                            Write-Log "  [OK] Gekopieerd van '$basedOn': $($grp.Name)" "SUCCESS"
                            $groupsOK++
                        } catch {
                            Write-Log "  [FAIL] AD groep '$($grp.Name)' mislukt: $($_.Exception.Message)" "ERROR"
                            $groupsFailed++
                        }
                    }
                    Write-Log "=== Kopie resultaat: $groupsOK OK, $groupsFailed mislukt ==="
                } else {
                    Write-Log "Brongebruiker '$basedOn' niet gevonden in AD" "ERROR"
                }
            } catch { Write-Log "Groepen kopieren mislukt: $($_.Exception.Message)" "ERROR" }
        }

        # Handmatig geselecteerde AD groepen toevoegen
        if ($created -and $manualADGroups.Count -gt 0) {
            Write-Log "=== Handmatige AD groepen toevoegen ($($manualADGroups.Count) geselecteerd) ==="
            $manOK = 0; $manFailed = 0
            foreach ($grpName in $manualADGroups) {
                try {
                    Add-ADGroupMember -Identity $grpName -Members $loginname -Server $dcServer -ErrorAction Stop
                    $NewGroups += if ($NewGroups) { ", $grpName" } else { $grpName }
                    Write-Log "  [OK] Handmatig toegevoegd: $grpName" "SUCCESS"
                    $manOK++
                } catch {
                    $errMsg = $_.Exception.Message
                    if ($errMsg -like "*already a member*") {
                        Write-Log "  [SKIP] '$grpName' - al lid (via kopie)" "WARN"
                    } else {
                        Write-Log "  [FAIL] '$grpName' mislukt: $errMsg" "ERROR"
                        $manFailed++
                    }
                }
            }
            Write-Log "=== Handmatig resultaat: $manOK OK, $manFailed mislukt ==="
        }

        # =============================================
        # HYBRID FLOW: Graph → Sync → Wacht → 365 Inrichten
        # =============================================
        if ($created -and $global:StartSync) {
            Write-Log "=== START Hybrid 365 inrichting ==="

            # Stap 1: Graph login EERST (nodig voor polling en 365 inrichting)
            Write-Log "Graph modules laden voor hybrid inrichting..."
            $modulesOK = Install-GraphModules
            $graphOK = $false
            if ($modulesOK) {
                # Gebruik AD domein als TenantId om device registration prompt te voorkomen
                $tenantDomain = ""
                # Gebruik het e-maildomein (UPN suffix) - NIET het AD domein (.local werkt niet)
                $tenantDomain = ($emailaddress -split '@')[1]
                # Fallback: eerste UPN suffix uit forest
                if (-not $tenantDomain -or $tenantDomain -like "*.local") {
                    try {
                        $forest = Get-ADForest -ErrorAction SilentlyContinue
                        if ($forest.UPNSuffixes -and $forest.UPNSuffixes.Count -gt 0) {
                            $tenantDomain = $forest.UPNSuffixes[0]
                        }
                    } catch {}
                }
                # Laatste fallback: AD domein (alleen als het geen .local is)
                if (-not $tenantDomain -or $tenantDomain -like "*.local") {
                    $adRoot = (Get-ADDomain -ErrorAction SilentlyContinue).DnsRoot
                    if ($adRoot -and $adRoot -notlike "*.local") { $tenantDomain = $adRoot }
                }
                # Als alles .local is: geen TenantId meegeven (standaard login)
                if ($tenantDomain -like "*.local") {
                    Write-Log "Geen geldig tenant domein gevonden (.local) - standaard Graph login" "WARN"
                    $tenantDomain = ""
                }
                Write-Log "Tenant domein: $(if ($tenantDomain) { $tenantDomain } else { '(standaard)' })"
                if ($tenantDomain) {
                    $graphOK = Connect-GraphInteractive -TenantId $tenantDomain
                } else {
                    $graphOK = Connect-GraphInteractive
                }
            } else {
                Write-Log "Graph modules niet beschikbaar - 365 inrichting overgeslagen" "WARN"
                Write-Log "Installeer handmatig: Install-Module Microsoft.Graph -Scope AllUsers -Force" "WARN"
            }
            if ($graphOK) {
                Write-Log "Graph verbonden als: $((Get-MgContext).Account)" "SUCCESS"
                $graphAccount = (Get-MgContext).Account
                if ($graphAccount -like "*#EXT#*") {
                    Write-Log "LET OP: Ingelogd als gastaccount - licenties en bepaalde bewerkingen zijn mogelijk beperkt" "WARN"
                    Write-Log "Gebruik een admin account dat direct lid is van de tenant voor volledige functionaliteit" "WARN"
                }

                # Stap 2: Start AD Sync en wacht tot user in 365 verschijnt
                $userInCloud = Start-ADSyncAndWait -UserPrincipalName $emailaddress -DisplayName $displayname -MaxWaitMinutes 10 -PollIntervalSeconds 15

                if ($userInCloud) {
                    $cloudUser = Get-MgUser -Filter "UserPrincipalName eq '$emailaddress'" -ErrorAction SilentlyContinue
                    # Fallback: zoek op mail of displayname
                    if (-not $cloudUser) {
                        Write-Log "UPN filter match niet, zoeken op mail..." "WARN"
                        $cloudUser = Get-MgUser -Filter "mail eq '$emailaddress'" -ErrorAction SilentlyContinue
                    }
                    if (-not $cloudUser) {
                        Write-Log "Mail filter match niet, zoeken op displayname..." "WARN"
                        $cloudUser = Get-MgUser -Filter "displayName eq '$displayname'" -ErrorAction SilentlyContinue | Select-Object -First 1
                    }

                    if ($cloudUser) {
                        Write-Log "Cloud user gevonden: $($cloudUser.UserPrincipalName) (Id: $($cloudUser.Id))" "SUCCESS"

                        # Stap 3: UsageLocation instellen
                        Write-Log "UsageLocation instellen: $usageLocation"
                        try {
                            Update-MgUser -UserId $cloudUser.Id -UsageLocation $usageLocation -ErrorAction Stop
                            Write-Log "UsageLocation '$usageLocation' ingesteld" "SUCCESS"
                        } catch {
                            Write-Log "UsageLocation instellen mislukt: $($_.Exception.Message)" "ERROR"
                        }

                        # Stap 4: Licentie selectie dialog (met echte tenant licenties)
                        $licenseAssigned = $false
                        Write-Log "Licenties ophalen uit tenant..."
                        try {
                            $global:TenantLicenses = Get-MgSubscribedSKU -All -Property @("SkuId","SkuPartNumber","ConsumedUnits","PrepaidUnits") -ErrorAction Stop |
                                Select-Object *, @{Name = "ActiveUnits"; Expression = { ($_ | Select-Object -ExpandProperty PrepaidUnits).Enabled } } |
                                Select-Object SkuId, SkuPartNumber, ActiveUnits, ConsumedUnits |
                                Where-Object { $_.ActiveUnits -gt 0 }
                            Write-Log "Licenties opgehaald: $($global:TenantLicenses.Count) typen"
                        } catch { Write-Log "Licenties ophalen mislukt: $($_.Exception.Message)" "ERROR" }

                        if ($global:TenantLicenses -and $global:TenantLicenses.Count -gt 0) {
                            # Licentie selectie popup
                            $licForm = New-Object System.Windows.Forms.Form
                            $licForm.Text = "365 Licentie toewijzen - $emailaddress"
                            $licForm.Size = New-Object System.Drawing.Size(500, 200)
                            $licForm.StartPosition = "CenterScreen"; $licForm.FormBorderStyle = 'FixedSingle'
                            $licForm.MaximizeBox = $false; $licForm.MinimizeBox = $false; $licForm.TopMost = $true
                            $licForm.BackColor = $BackColor; $licForm.ForeColor = $ForeColor; $licForm.Font = $Form_Font

                            $lblLic = New-Object System.Windows.Forms.Label
                            $lblLic.Text = "Selecteer licentie(s) voor $displayname :"
                            $lblLic.AutoSize = $true; $lblLic.Location = New-Object System.Drawing.Point(20, 20)
                            [void]$licForm.Controls.Add($lblLic)

                            $chkLic = New-Object System.Windows.Forms.CheckedListBox
                            $chkLic.Location = New-Object System.Drawing.Point(20, 50); $chkLic.Width = 440; $chkLic.Height = 140
                            $chkLic.BackColor = $TextBoxBackColor; $chkLic.ForeColor = $ForeColor
                            $chkLic.Font = $TextBoxFont; $chkLic.CheckOnClick = $true
                            $chkLic.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
                            foreach ($lic in ($global:TenantLicenses | Sort-Object SkuPartNumber)) {
                                $avail = [int]$lic.ActiveUnits - [int]$lic.ConsumedUnits
                                [void]$chkLic.Items.Add("$($lic.SkuPartNumber) ($avail van $($lic.ActiveUnits) beschikbaar)")
                            }
                            [void]$licForm.Controls.Add($chkLic)

                            $licForm.Height = 280

                            $btnLicOK = New-Object System.Windows.Forms.Button
                            $btnLicOK.Text = "Toewijzen"; $btnLicOK.Size = New-Object System.Drawing.Size(120, 30)
                            $btnLicOK.Location = New-Object System.Drawing.Point(340, 210)
                            $btnLicOK.FlatStyle = "Flat"; $btnLicOK.BackColor = $ButtonColor; $btnLicOK.ForeColor = $ForeColor
                            $btnLicOK.Font = $Font10B
                            $btnLicOK.FlatAppearance.BorderColor = $ButtonBorderColor; $btnLicOK.FlatAppearance.BorderSize = 2
                            $btnLicOK.Add_Click({
                                $selected = @()
                                for ($i = 0; $i -lt $chkLic.Items.Count; $i++) {
                                    if ($chkLic.GetItemChecked($i)) { $selected += $chkLic.Items[$i].ToString() }
                                }
                                $licForm.Tag = $selected
                                $licForm.Close()
                            })
                            [void]$licForm.Controls.Add($btnLicOK)

                            $btnLicSkip = New-Object System.Windows.Forms.Button
                            $btnLicSkip.Text = "Overslaan"; $btnLicSkip.Size = New-Object System.Drawing.Size(100, 30)
                            $btnLicSkip.Location = New-Object System.Drawing.Point(210, 210)
                            $btnLicSkip.FlatStyle = "Flat"; $btnLicSkip.BackColor = $BackColor; $btnLicSkip.ForeColor = $ForeColor
                            $btnLicSkip.Font = $Font9
                            $btnLicSkip.FlatAppearance.BorderColor = [System.Drawing.Color]::Gray; $btnLicSkip.FlatAppearance.BorderSize = 1
                            $btnLicSkip.Add_Click({ $licForm.Tag = $null; $licForm.Close() })
                            [void]$licForm.Controls.Add($btnLicSkip)

                            [void]$licForm.ShowDialog()

                            $selectedLics = $licForm.Tag
                            [void]$licForm.Dispose()

                            if ($selectedLics -and $selectedLics.Count -gt 0) {
                                foreach ($selectedLic in $selectedLics) {
                                    $skuPart = ($selectedLic -split " \(")[0].Trim()
                                    $sku = $global:TenantLicenses | Where-Object { $_.SkuPartNumber -eq $skuPart }
                                    if ($sku) {
                                        Write-Log "Licentie toewijzen: $skuPart (SkuId: $($sku.SkuId))"
                                        try {
                                            Set-MgUserLicense -UserId $cloudUser.Id -BodyParameter @{ AddLicenses = @(@{ DisabledPlans = @(); SkuId = $sku.SkuId }); RemoveLicenses = @() } -ErrorAction Stop
                                            Write-Log "Licentie '$skuPart' toegewezen" "SUCCESS"
                                            $licenseAssigned = $true
                                        } catch {
                                            Write-Log "Licentie toewijzen mislukt: $($_.Exception.Message)" "ERROR"
                                        }
                                    }
                                }
                            } else {
                                Write-Log "Geen licentie geselecteerd - overgeslagen"
                            }
                        } else {
                            Write-Log "Geen licenties gevonden in tenant - licentie dialog overgeslagen" "WARN"
                            $ctx = Get-MgContext -ErrorAction SilentlyContinue
                            if ($ctx -and $ctx.Account -like "*#EXT#*") {
                                Write-Log "Je bent ingelogd als gastaccount ($($ctx.Account)) - dit account heeft mogelijk geen rechten om licenties te lezen" "WARN"
                                Write-Log "Log in met een account dat direct lid is van de tenant (geen gastaccount)" "WARN"
                            }
                        }

                        # Stap 5: 365 groepen kopieren van BasedOn gebruiker
                        if ($basedOn) {
                            $copyGroups365 = [System.Windows.Forms.MessageBox]::Show(
                                "Wil je de 365 groepen van '$basedOn' kopieren naar $displayname ?",
                                "365 Groepen kopieren",
                                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                                [System.Windows.Forms.MessageBoxIcon]::Question
                            )
                            if ($copyGroups365 -eq [System.Windows.Forms.DialogResult]::Yes) {
                            Write-Log "=== 365 groepen kopieren van '$basedOn' ==="
                            $cloud365OK = 0; $cloud365Failed = 0; $cloud365Skipped = 0
                            try {
                                $srcCloudUser = Get-MgUser -Filter "DisplayName eq '$basedOn'" -ErrorAction SilentlyContinue | Select-Object -First 1
                                if (-not $srcCloudUser) {
                                    $srcADUser = $global:ADusers | Where-Object { $_.Name -eq $basedOn }
                                    if ($srcADUser) {
                                        $srcCloudUser = Get-MgUser -Filter "UserPrincipalName eq '$($srcADUser.UserPrincipalName)'" -ErrorAction SilentlyContinue
                                    }
                                }
                                if ($srcCloudUser) {
                                    Write-Log "Brongebruiker in 365: $($srcCloudUser.DisplayName) (Id: $($srcCloudUser.Id))"
                                    $cloudGrps = Get-MgUserMemberOf -UserId $srcCloudUser.Id -All
                                    Write-Log "Brongebruiker is lid van $($cloudGrps.Count) 365 groepen/rollen"

                                    foreach ($g in $cloudGrps) {
                                        if ($g.AdditionalProperties.'@odata.type' -ne '#microsoft.graph.group') {
                                            $cloud365Skipped++; continue
                                        }
                                        $groupName = $g.AdditionalProperties.displayName
                                        $onPremSync = $g.AdditionalProperties.onPremisesSyncEnabled
                                        $secEnabled = $g.AdditionalProperties.securityEnabled
                                        $mailEnabled = $g.AdditionalProperties.mailEnabled
                                        if ($onPremSync -eq $true -and $secEnabled -eq $true -and $mailEnabled -ne $true) {
                                            Write-Log "  [SKIP] '$groupName' - on-prem synced (al via AD)"
                                            $cloud365Skipped++; continue
                                        }
                                        try {
                                            New-MgGroupMember -GroupId $g.Id -DirectoryObjectId $cloudUser.Id -ErrorAction Stop | Out-Null
                                            $NewGroups += if ($NewGroups) { ", $groupName" } else { $groupName }
                                            Write-Log "  [OK] 365 groep: $groupName" "SUCCESS"
                                            $cloud365OK++
                                        } catch {
                                            $errMsg = $_.Exception.Message
                                            if ($errMsg -like "*already exist*") {
                                                Write-Log "  [SKIP] '$groupName' - al lid" "WARN"
                                                $cloud365Skipped++
                                            } else {
                                                Write-Log "  [FAIL] '$groupName': $errMsg" "ERROR"
                                                $cloud365Failed++
                                            }
                                        }
                                    }
                                    Write-Log "=== 365 groepen: $cloud365OK OK, $cloud365Failed mislukt, $cloud365Skipped overgeslagen ==="
                                } else {
                                    Write-Log "Brongebruiker '$basedOn' niet gevonden in 365" "WARN"
                                }
                            } catch { Write-Log "365 groepen mislukt: $($_.Exception.Message)" "ERROR" }
                            } else {
                                Write-Log "365 groepen kopieren overgeslagen door gebruiker"
                            }
                        }


                        # Stap 6: 365 Inrichtingsformulier - handmatige groepen, DLs en shared mailboxes
                        Write-Log "365 inrichtingsformulier tonen..."
                        $selections365 = Show-365ConfigForm -UserDisplayName $displayname -UserEmail $emailaddress -UserId $cloudUser.Id
                        if ($selections365) {
                            $cfgResult = Process-365Selections -Selections $selections365 -UserId $cloudUser.Id -UserEmail $emailaddress -ExoAvailable $selections365.ExoAvailable -LicenseAssigned $licenseAssigned
                            if ($cfgResult) {
                                if ($cfgResult.OK -gt 0) {
                                    Write-Log "365 handmatige inrichting: $($cfgResult.OK) items toegevoegd" "SUCCESS"
                                }
                            }
                        }

                        # Stap 7: Shared mailbox rechten kopieren van BasedOn gebruiker
                        if ($basedOn -and $licenseAssigned -and $global:ExoConnected) {
                            Write-Log "Shared mailbox rechten ophalen van '$basedOn'..."
                            try {
                                $srcEmail = ""
                                $srcADUser = $global:ADusers | Where-Object { $_.Name -eq $basedOn }
                                if ($srcADUser) { $srcEmail = $srcADUser.UserPrincipalName }
                                if (-not $srcEmail -and $srcCloudUser) { $srcEmail = $srcCloudUser.UserPrincipalName }

                                if ($srcEmail) {
                                    $sharedMBs = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction SilentlyContinue
                                    $srcPermissions = @()
                                    foreach ($smb in $sharedMBs) {
                                        $perms = Get-MailboxPermission -Identity $smb.Identity -ErrorAction SilentlyContinue |
                                            Where-Object { $_.User -like "*$srcEmail*" -and $_.AccessRights -contains "FullAccess" -and -not $_.IsInherited }
                                        if ($perms) { $srcPermissions += $smb }
                                    }

                                    if ($srcPermissions.Count -gt 0) {
                                        $mbNames = ($srcPermissions | ForEach-Object { $_.DisplayName }) -join "`n"
                                        Write-Log "Brongebruiker heeft FullAccess op $($srcPermissions.Count) shared mailbox(es)"

                                        $copyMB = [System.Windows.Forms.MessageBox]::Show(
                                            "Brongebruiker '$basedOn' heeft toegang tot $($srcPermissions.Count) shared mailbox(es):`n`n$mbNames`n`nWil je deze rechten kopieren naar $displayname ?",
                                            "Shared Mailbox rechten kopieren",
                                            [System.Windows.Forms.MessageBoxButtons]::YesNo,
                                            [System.Windows.Forms.MessageBoxIcon]::Question
                                        )
                                        if ($copyMB -eq [System.Windows.Forms.DialogResult]::Yes) {
                                            Write-Log "Shared mailbox rechten kopieren..."
                                            $mbxReady = Wait-ForMailbox -UserEmail $emailaddress -MaxWaitMinutes 5 -PollIntervalSeconds 15
                                            if ($mbxReady) {
                                                foreach ($smb in $srcPermissions) {
                                                    try {
                                                        Add-MailboxPermission -Identity $smb.Identity -User $emailaddress -AccessRights FullAccess -AutoMapping $true -ErrorAction Stop | Out-Null
                                                        Write-Log "  [OK] $($smb.DisplayName) - FullAccess" "SUCCESS"
                                                    } catch {
                                                        if ($_.Exception.Message -like "*already*") {
                                                            Write-Log "  [SKIP] $($smb.DisplayName) - bestaat al" "WARN"
                                                        } else {
                                                            Write-Log "  [FAIL] $($smb.DisplayName) FullAccess: $($_.Exception.Message)" "ERROR"
                                                        }
                                                    }
                                                    try {
                                                        Add-RecipientPermission -Identity $smb.Identity -Trustee $emailaddress -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                                                        Write-Log "  [OK] $($smb.DisplayName) - SendAs" "SUCCESS"
                                                    } catch {
                                                        if ($_.Exception.Message -like "*already*") { } else {
                                                            Write-Log "  [WARN] $($smb.DisplayName) SendAs: $($_.Exception.Message)" "WARN"
                                                        }
                                                    }
                                                }
                                            } else {
                                                Write-Log "Mailbox niet beschikbaar na wachten - shared mailbox rechten overgeslagen" "WARN"
                                            }
                                        } else {
                                            Write-Log "Shared mailbox rechten kopieren overgeslagen door gebruiker"
                                        }
                                    } else {
                                        Write-Log "Brongebruiker heeft geen shared mailbox rechten"
                                    }
                                }
                            } catch { Write-Log "Shared mailbox rechten ophalen mislukt: $($_.Exception.Message)" "WARN" }
                        } elseif ($basedOn -and $licenseAssigned -and -not $global:ExoConnected) {
                            Write-Log "EXO niet verbonden - shared mailbox rechten niet gekopieerd" "WARN"
                        }
                    } else {
                        Write-Log "Cloud user $emailaddress niet gevonden via UPN/mail/displayname" "ERROR"
                    }
                } else {
                    Write-Log "User niet in 365 verschenen na sync - 365 inrichting overgeslagen" "WARN"
                    Write-Log "Controleer AD Sync status en richt de user handmatig in" "WARN"
                }
            } else {
                Write-Log "Graph login mislukt - 365 inrichting overgeslagen" "WARN"
                Write-Log "Richt de user handmatig in via het 365 Admin Portal" "WARN"
            }
            Write-Log "=== EINDE Hybrid 365 inrichting ==="
        }
    }
    Write-Log "=== EINDE On-Prem verwerking ==="
}

# ============================================================================
# STAP 7: Resultaat tonen
# ============================================================================
Write-Log "========================================"

if ($errorOccured) {
    Write-Log "RESULTAAT: MISLUKT - $errorMessage" "ERROR"
    [System.Windows.Forms.MessageBox]::Show($errorMessage, "Fout", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
} else {
    Write-Log "RESULTAAT: SUCCES" "SUCCESS"
    Write-Log "  Email: $emailaddress"
    Write-Log "  Weergavenaam: $displayname"
    Write-Log "  Locatie: $usageLocation"
    if ($department) { Write-Log "  Afdeling: $department" }
    if ($title) { Write-Log "  Functie: $title" }
    if ($manager) { Write-Log "  Manager: $manager" }
    if ($NewGroups) { Write-Log "  Groepen: $NewGroups" }

    $resultMsg = "Gebruiker succesvol aangemaakt!`n`n"
    $resultMsg += "Email: $emailaddress`n"
    $resultMsg += "Wachtwoord: $password`n"
    $resultMsg += "Weergavenaam: $displayname`n"
    if ($department) { $resultMsg += "Afdeling: $department`n" }
    if ($title) { $resultMsg += "Functie: $title`n" }
    if ($manager) { $resultMsg += "Manager: $manager`n" }
    if ($NewGroups) { $resultMsg += "Groepen: $NewGroups`n" }
    if ($global:EnvironmentMode -eq "OnPrem") {
        if ($global:StartSync) {
            $resultMsg += "`nHybrid 365 inrichting uitgevoerd (zie logbestand voor details)."
        } else {
            $resultMsg += "`nLET OP: Azure AD Sync niet gestart - richt de user handmatig in via 365."
        }
    }

    $clipText = "$emailaddress   $password"
    try { [System.Windows.Forms.Clipboard]::SetText($clipText) } catch {}

    [System.Windows.Forms.MessageBox]::Show("$resultMsg`n`nEmail + wachtwoord zijn naar het klembord gekopieerd.", "Succes", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

Write-Log "========================================"
Write-Log "Logbestand: $LOGFile"

# ============================================================================
# Cleanup
# ============================================================================
try { Disconnect-MgGraph -ErrorAction Ignore > $null } catch {}
try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Ignore > $null } catch {}

[void]$main_form.Dispose()

try {
    Remove-Module -Name Microsoft.Graph.Authentication -Force -ErrorAction SilentlyContinue > $null
    Remove-Module -Name Microsoft.Graph.Users -Force -ErrorAction SilentlyContinue > $null
    Remove-Module -Name Microsoft.Graph.Users.Actions -Force -ErrorAction SilentlyContinue > $null
    Remove-Module -Name Microsoft.Graph.Groups -Force -ErrorAction SilentlyContinue > $null
    Remove-Module -Name Microsoft.Graph.Identity.DirectoryManagement -Force -ErrorAction SilentlyContinue > $null
    Remove-Module -Name ExchangeOnlineManagement -Force -ErrorAction SilentlyContinue > $null
    Remove-Module -Name ActiveDirectory -Force -ErrorAction SilentlyContinue > $null
} catch {}

Write-Log "Script beeindigd."
