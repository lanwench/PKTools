#requires -version 4
Function New-PKJargonIpsum {
<#
.SYNOPSIS
    Want a wall of mission statements? This function generates jargon-filled Lorem Ipsum text from internal dictionary arrays of words, 
    so you can fit in at your next meeting!

.DESCRIPTION
    Generates jargon-filled Lorem Ipsum text from internal dictionary arrays of words, so you can fit in at your next meeting!  
    All available words are stored within the function; no external connectivity is required
    It allows you to specify the number of sentences and paragraphs to generate, as well as the formatting options for 
    dividing paragraphs and the starting word type for sentences.

.NOTES
    Name    : Function_New-PKJargonIpsum.ps1
    Created : 2024-04-26
    Author  : Paula Kingsley
    Version : 01.00.0000
    History :
    
        ** PLEASE KEEP $VERSION UPDATED IN PROCESS BLOCK **

        v01.00.0000 - 2024-04-24\6 - Created script based on Tom Sherman's json data from link

.PARAMETER NumSentence
    Specifies the number of sentences to generate. Must be an even number between 2 and 100.

.PARAMETER NumParagraph
    Specifies the number of paragraphs to generate. Must be an even number between 2 and 20.

.PARAMETER Divider
    Specifies how paragraphs should be divided. Valid values are "SingleLine", "TwoLines", "Space", and "Emoji".

.PARAMETER FirstWord
    Specifies whether sentences should start with a verb or an adverb. Valid values are "Verb" and "Adverb".

.EXAMPLE
    PS C:\> New-PKJargonIpsum -NumSentence 6 -NumParagraph 4 -FirstWord Verb
    Generates 6 sentences divided into 4 paragraphs, with each paragraph separated by two lines. Sentences start with verbs.

.EXAMPLE
    PS C:\> New-PKJargonIpsum -NumSentence 10 -NumParagraph 2 -Divider Space 
    Generates 10 sentences divided into 2 paragraphs, with each paragraph separated by a space. Sentences start with adverbs.
#>
[Cmdletbinding()]
Param(
    [Parameter(
        Position = 0,
        HelpMessage = "Number of sentences to generate, even numbers between 2 and 100",
        ValueFromPipeline
    )]
    [ValidateRange(2,100)]
    [ValidateScript({if ($_ % 2 -eq 0) {$true} else {throw "Even numbers only please!"}})]
    [int]$NumSentence = 20,

    [Parameter(
        HelpMessage = "Number of paragraphs to generate, even numbers between 2 and 10"
    )]
    [ValidateRange(2,20)]
    [ValidateScript({if ($_ % 2 -eq 0) {$true} else {throw "Even numbers only please!"}})]
    [int]$NumParagraph = 4,

    [Parameter(
        HelpMessage = "Split up paragraphs with spaces, single line breaks, or double line breaks"
    )]
    [ValidateSet("SingleLine","TwoLines","Space","Emoji")]
    [string]$Divider = "TwoLines",

    [Parameter(
        HelpMessage = "Begin sentence with verb or with adverb (default is adverb)"
    )]
    [ValidateSet("Verb","Adverb")]
    [string]$FirstWord = "Adverb"
)
Begin {

    # Current version (please keep up to date from comment block)
    [version]$Version = "01.00.0000"

    # Show our settings
    $ScriptName = $MyInvocation.MyCommand.Name
    [switch]$PipelineInput = $MyInvocation.ExpectingInput
    $CurrentParams = $PSBoundParameters
    $MyInvocation.MyCommand.Parameters.keys | Where-Object { $CurrentParams.keys -notContains $_ } | 
    Where-Object { Test-Path variable:$_ } | ForEach-Object {
        $CurrentParams.Add($_, (Get-Variable $_).value)
    }
    $CurrentParams.Add("PipelineInput", $PipelineInput)
    $CurrentParams.Add("ScriptName", $ScriptName)
    $CurrentParams.Add("ScriptVersion", $Version)
    Write-Verbose "PSBoundParameters: `n`t$($CurrentParams | Format-Table -AutoSize | out-string )"

    #region Words! 
    $AllAdverbs = @("appropriately",
    "assertively",
    "authoritatively",
    "collaboratively",
    "compellingly",
    "competently",
    "completely",
    "continually",
    "conveniently",
    "credibly",
    "distinctively",
    "dramatically",
    "dynamically",
    "efficiently",
    "energistically",
    "enthusiastically",
    "globally",
    "holistically",
    "interactively",
    "intrinsically",
    "monotonectally",
    "objectively",
    "performantly",
    "phosfluorescently",
    "proactively",
    "professionally",
    "progressively",
    "quickly",
    "rapidiously",
    "seamlessly",
    "synergistically",
    "uniquely",
    "fungibly")

    $Allverbs = @("actualize",
    "administrate",
    "aggregate",
    "architect",
    "benchmark",
    "brand",
    "build",
    "communicate",
    "conceptualize",
    "coordinate",
    "create",
    "cultivate",
    "customize",
    "deliver",
    "deploy",
    "develop",
    "disintermediate",
    "disseminate",
    "drive",
    "embrace",
    "e-enable",
    "embiggen",
    "empower",
    "enable",
    "engage",
    "engineer",
    "enhance",
    "envisioneer",
    "evisculate",
    "evolve",
    "expedite",
    "exploit",
    "extend",
    "fabricate",
    "facilitate",
    "fashion",
    "formulate",
    "foster",
    "generate",
    "grow",
    "harness",
    "impact",
    "implement",
    "incentivize",
    "incubate",
    "initiate",
    "innovate",
    "integrate",
    "iterate",
    "leverage existing",
    "leverage other's",
    "maintain",
    "matrix",
    "maximize",
    "mesh",
    "monetize",
    "morph",
    "myocardinate",
    "negotiate",
    "network",
    "optimize",
    "orchestrate",
    "parallel task",
    "plagiarize",
    "pontificate",
    "predominate",
    "procrastinate",
    "productivate",
    "productize",
    "promote",
    "provide access to",
    "pursue",
    "recaptiualize",
    "reconceptualize",
    "redefine",
    "re-engineer",
    "reintermediate",
    "reinvent",
    "repurpose",
    "restore",
    "revolutionize",
    "scale",
    "seize",
    "simplify",
    "strategize",
    "streamline",
    "supply",
    "syndicate",
    "synergize",
    "synthesize",
    "target",
    "transform",
    "transition",
    "underwhelm",
    "unleash",
    "utilize",
    "visualize",
    "whiteboard",
    "cloudify",
    "right-shore")

    $AllAdjectives = @("24/7",
    "24/365",
    "accurate",
    "adaptive",
    "alternative",
    "an expanded array of",
    "B2B",
    "B2C",
    "backend",
    "backward-compatible",
    "best-of-breed",
    "bleeding-edge",
    "bricks-and-clicks",
    "business",
    "clicks-and-mortar",
    "client-based",
    "client-centered",
    "client-centric",
    "client-focused",
    "collaborative",
    "compelling",
    "competitive",
    "cooperative",
    "corporate",
    "cost effective",
    "covalent",
    "cromulent"
    "cross-functional",
    "cross-media",
    "cross-platform",
    "cross-unit",
    "customer-directed",
    "customized",
    "cutting-edge",
    "distinctive",
    "distributed",
    "diverse",
    "dynamic",
    "e-business",
    "economically-sound",
    "effective",
    "efficient",
    "emerging",
    "empowered",
    "enabled",
    "end-to-end",
    "enterprise",
    "enterprise-wide",
    "equity-invested",
    "error-free",
    "ethical",
    "excellent",
    "exceptional",
    "extensible",
    "extensive",
    "flexible",
    "focused",
    "frictionless",
    "front-end",
    "fully-researched",
    "fully-tested",
    "functional",
    "functionalized",
    "future-proof",
    "global",
    "go-forward",
    "goal-oriented",
    "granular",
    "high standards in",
    "high-payoff",
    "high-quality",
    "highly-efficient",
    "highly-performant",
    "holistic",
    "impactful",
    "inexpensive",
    "innovative",
    "installed base",
    "integrated",
    "interactive",
    "interdependent",
    "intermandated",
    "interoperable",
    "intuitive",
    "just-in-time",
    "leading-edge",
    "leveraged",
    "long-term, high-impact",
    "low-risk, high-yield",
    "magnetic",
    "maintainable",
    "market positioning",
    "market-driven",
    "mission-critical",
    "multidisciplinary",
    "multifunctional",
    "multimedia-based",
    "next-generation",
    "one-to-one",
    "open-source",
    "optimal",
    "orthogonal",
    "out-of-the-box",
    "pandemic",
    "parallel",
    "performance-based",
    "performant,"
    "plug-and-play",
    "premier",
    "premium",
    "principle-centered",
    "proactive",
    "process-centric",
    "professional",
    "progressive",
    "prospective",
    "quality",
    "real-time",
    "reliable",
    "resource-sucking",
    "resource-maximizing",
    "resource-leveling",
    "revolutionary",
    "robust",
    "scalable",
    "seamless",
    "standalone",
    "standardized",
    "standards-compliant",
    "state-of-the-art",
    "sticky",
    "strategic",
    "superior",
    "sustainable",
    "synergistic",
    "tactical",
    "team-building",
    "team-driven",
    "technically-sound",
    "timely",
    "top-line",
    "transparent",
    "turnkey",
    "ubiquitous",
    "unique",
    "user-centric",
    "user-friendly",
    "value-added",
    "vertical",
    "viral",
    "virtual",
    "visionary",
    "web-enabled",
    "wireless",
    "world-class",
    "worldwide",
    "fungible",
    "cloud-ready",
    "elastic",
    "hyper-scale",
    "on-demand",
    "cloud-based",
    "cloud-centric",
    "cloudified",
    "agile",
    "omni-channel")

    $AllNouns = @("action items",
    "alignments",
    "applications",
    "architectures",
    "bandwidth",
    "benefits",
    "best practices",
    "catalysts for change",
    "channels",
    "collaboration and idea-sharing",
    "communities",
    "content",
    "convergence",
    "core competencies",
    "customer service",
    "data",
    "deliverables",
    "e-business",
    "e-commerce",
    "e-markets",
    "e-tailers",
    "e-services",
    "experiences",
    "expertise",
    "functionalities",
    "growth strategies",
    "human capital",
    "ideas",
    "imperatives",
    "infomediaries",
    "information",
    "infrastructures",
    "initiatives",
    "innovation",
    "intellectual capital",
    "interfaces",
    "internal or organic sources",
    "leadership",
    "leadership skills",
    "manufactured products",
    "markets",
    "materials",
    "meta-services",
    "methodologies",
    "methods of empowerment",
    "metrics",
    "mindshare",
    "models",
    "networks",
    "niches",
    "niche markets",
    "opportunities",
    "outside-the-box thinking",
    "outsourcing",
    "paradigms",
    "partnerships",
    "platforms",
    "portals",
    "potentialities",
    "process improvements",
    "processes",
    "products",
    "quality vectors",
    "relationships",
    "resources",
    "results",
    "ROI",
    "scenarios",
    "schemas",
    "services",
    "solutions",
    "sources",
    "strategic theme areas",
    "supply chains",
    "synergy",
    "systems",
    "technologies",
    "technology",
    "testing procedures",
    "total linkage",
    "users",
    "value",
    "vortals",
    "web-readiness",
    "web services",
    "fungibility",
    "clouds",
    "nosql",
    "storage",
    "virtualization",
    "scrums",
    "sprints",
    "wins",
    "blue-sky thinking")

    $StatementVerbs = @(
        "achieve",
        "deliver",
        "meet",
        "realize",
        "create",
        "fulfill",
        "attain",
        "accomplish",
        "envision"
    )

    $StatementAdjectives = @(
        "steadfast",
        "resolute",
        "firm",
        "unyielding",
        "unflinching",
        "determined",
        "staunch",
        "steadfast",
        "unswerving",
        "unfaltering",
        "unshakeable",
        "unrelenting",
        "persistent",
        "unflagging",
        "unhesitating",
        "unbending",
        "relentless",
        "persevering"
    )

    $StatementNouns = @(
        "commitment",
        "dedication",
        "devotion",
        "loyalty",
        "fidelity",
        "allegiance",
        "obligation",
        "responsibility",
        "duty",
        "accountability"
    )
        
    $StatementOutcomes = @(
        "excellence",
        "customer delight",
        "innovation",
        "culture",
        "quality",
        "sustainability",
        "growth",
        "customer success"
        "customer profitability"
    )

    <# Future use?
    $StatementGoals = @(
        "our mission is",
        "the primary goal is",
        "the company's primary objective is",
        "our collective strategic purpose is",
        "our aim is",
        "our corporate vision is"
    )
    #>

    $StatementConnectors = @(
        "while remaining laser-focused on",
        "whilst continuing to",
        "without failing to",
        "and keep a clear focus on our goal to",
        "and never forget to",
        "while always remembering to",
        "but never forget to",
        "while never failing to"
        "without losing sight of our goal to",
        "whilst maintaining a clear path towards our vision to"
    )

    $StatementActions = @(
        "we promise to"
        "our principal tenet is to",
        "we commit to",
        "we cannot and will not rest until we" 
        "we"
        "our organization will"
        "our solemn vow is to"
    )

    $StatementOpenings = @(
        "As part of our $($StatementAdjectives | Get-Random -Count 1) $($StatementNouns | Get-Random -Count 1) to $($StatementOutcomes | Get-Random -Count 1), $($StatementActions | Get-Random -Count 1)",
        "So that we can $($StatementVerbs | Get-Random -Count 1) $($StatementOutcomes | Get-Random -Count 1), $($StatementAction | Get-Random -Count 1)",
        "In order to $($AllAdverbs | Get-Random -Count 1) $($StatementVerbs | Get-Random -Count 1) $($StatementOutcomes | Get-Random -Count 1),  $($StatementActions | Get-Random -Count 1)",
        "To ensure we $($StatementVerbs | Get-Random -Count 1) our goal to $($StatementOutcomes | Get-Random -Count 1), $($StatementActions | Get-Random -Count 1)",
        "As the most critical component of our $($StatementAdjectives | Get-Random -Count 1) focus on $($StatementOutcomes | Get-Random -Count 1), $($StatementActions | Get-Random -Count 1)",
        "With our unwavering commitment to $($StatementOutcomes | Get-Random -Count 1), $($StatementActions | Get-Random -Count 1)",
        "Driven by our $($StatementAdjectives | Get-Random -Count 1) $($StatementNouns | Get-Random -Count 1), we aim to $($StatementVerbs | Get-Random -Count 1) $($StatementOutcomes | Get-Random -Count 1)",
        "In the spirit of $($StatementAdjectives | Get-Random -Count 1) $($StatementNouns | Get-Random -Count 1), we strive to $($StatementVerbs | Get-Random -Count 1) $($StatementOutcomes | Get-Random -Count 1)",
        "Guided by our $($StatementAdjectives | Get-Random -Count 1) commitment to $($StatementOutcomes | Get-Random -Count 1), $($StatementActions | Get-Random -Count 1)",
        "In pursuit of $($StatementOutcomes | Get-Random -Count 1), we $($StatementActions | Get-Random -Count 1)",
        "Our $($StatementAdjectives | Get-Random -Count 1) dedication to $($StatementOutcomes | Get-Random -Count 1) compels us to $($StatementActions | Get-Random -Count 1)",
        "To uphold our $($StatementAdjectives | Get-Random -Count 1) $($StatementNouns | Get-Random -Count 1), we pledge to $($StatementActions | Get-Random -Count 1)",
        "In alignment with our goal to $($StatementVerbs | Get-Random -Count 1) $($StatementOutcomes | Get-Random -Count 1), $($StatementActions | Get-Random -Count 1)"
    )

    #For title case
    #$textInfo = (Get-Culture).TextInfo

    #endregion Words! 

    #region Functions, etc
    Function _GetEmoji{
        $start = 0x1F600  # Start of the Unicode character range for emojis
        $end = 0x1F64F  # End of the range
        # Generate the list of emoji Unicode characters, then convert the integer to a char & output a unique one
        $Emojis = for ($i = $start; $i -le $end; $i++) {
            [char]::ConvertFromUtf32($i)
        }
        $Emojis  | Get-Random -Count 1
    }
    
    $DividerChar = Switch ($Divider) {
        SingleLine {"`n"}
        TwoLines {"`n`n"}
        Space {" "}
        Emoji {" $(_GetEmoji) "}
    }

    #endregion Functions, etc

    $Msg = "Generating $NumParagraph paragraphs of $NumSentence sentences of utter nonsense"
    Write-Verbose "[BEGIN: $ScriptName] $Msg"

}
Process {

    # Collect all the paragraphs here
    $ParagraphArr = @()
    $CurrentParaCount = 0

    # Do this for the total number of paragraphs
    For ($v = 0; $v -lt $NumParagraph; $v++) {

        $CurrentParaCount++
        $Msg = "Creating $NumSentence sentences...."
        Write-Verbose "[Paragraph $CurrentParaCount/$NumParagraph] $Msg"

        # Get sentences
        $SentenceArr = @()
        For ($w = 0; $w -lt $NumSentence; $w++) {
            $Noun = $AllNouns | Get-Random -Count 1
            $Verb = $AllVerbs | Get-Random -Count 1
            $Adverb = $AllAdverbs | Get-Random -Count 1
            $Adjective = $AllAdjectives | Get-Random -Count 1
            $SentenceArr += Switch ($FirstWord) {
                "Verb" { "$Verb $Adjective $Noun $Adverb" }
                "Adverb" {"$Adverb $Verb $Adjective $Noun" }
            }
        }    

        # We want to combine every 2 of these sentences
        $CombinedSentences = for ($x = 0; $x -lt $SentenceArr.Length; $x += 2) { 
            "$($SentenceArr[$x]) $($StatementConnectors | Get-Random -Count 1) $($SentenceArr[$x + 1])"
        }
        
        # Then we prepend with a statement opening
        $Statements = Foreach ($y in $CombinedSentences)  {
            "$($StatementOpenings | Get-Random -Count 1) $y."
        } 
        
        # We don't want things to sort too obviously, so we'll shuffle them, although it may not work
        $ParagraphArr += ($Statements | Get-Random -Count $($Statements.Count))
    }

    # Now we want to combine every 5 of these paragraphs
    $Msg = "Combining every 5 paragraphs into one"
    $FinalParagraphs = @()
    for ($z = 0; $z -lt $ParagraphArr.Length; $z += 5) {
        $FinalParagraphs += (($ParagraphArr | Get-Random -Count $ParagraphArr.Length)[$z..($z+4)] -join " ")
    }

    # Output the final paragraphs split up as we chose
    ($FinalParagraphs | Get-Random -Count $FinalParagraphs.Length) -join $DividerChar

}
End {
    Write-Verbose "[END: $ScriptName]"
}
} #end function

