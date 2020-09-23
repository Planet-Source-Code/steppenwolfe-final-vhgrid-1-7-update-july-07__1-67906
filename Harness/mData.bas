Attribute VB_Name = "mData"
Option Explicit

Public m_sAppName()     As String
Public m_sAppDesc()     As String
Public m_sAppStats()    As String
Public m_sAppData()     As String

Public m_sMedia()       As String
Public m_sLyrics()      As String
Public m_sTitles()      As String
Public m_sDesc()        As String

Public Sub CreateAppData()

    ReDim m_sAppName(4)
    ReDim m_sAppDesc(4)
    ReDim m_sAppStats(4)
    ReDim m_sAppData(4)
    
    m_sAppName(0) = "MicroSoft Access"
    m_sAppDesc(0) = "Gadi's Internet Threats"
    m_sAppStats(0) = "Size 7.1GB" & vbNewLine & "Relevance: None" & vbNewLine & "Active: No"
    m_sAppData(0) = "Various post-fix threat declarations, potential incidents, (and paranoia), recorded via forum archives. "

    m_sAppName(1) = "MicroSoft Excel"
    m_sAppDesc(1) = "Quarterly Breakdown"
    m_sAppStats(1) = "Size 27MB" & vbNewLine & "Active: No" & vbNewLine & "Archived: Yes"
    m_sAppData(1) = "A review of quarterly financial statements. Billing, and resolved account data to quarter ending August 05."
    
    m_sAppName(2) = "MicroSoft FrontPage"
    m_sAppDesc(2) = "New Home Page - OurSpace.com"
    m_sAppStats(2) = "Size: 17k" & vbNewLine & "Scripts: No" & vbNewLine & "Flash: No"
    m_sAppData(2) = "Welcome to OurSpace.com, where our dreams of a free internet come to life.. Todays Blogs: Ernest talks trash - Is Google taking over the Internet?"
    
    m_sAppName(3) = "MicroSoft PowerPoint"
    m_sAppDesc(3) = "INS Presentation: Tuesday April 7"
    m_sAppStats(3) = "Size 3.7MB" & vbNewLine & "Active: No" & vbNewLine & "Archived: Yes"
    m_sAppData(3) = "Securing edge devices without the overhead, Secure NMS, Snort updates, and using Isolation LANs to segment assets"

    m_sAppName(4) = "MicroSoft Word"
    m_sAppDesc(4) = "Meeting Notes: Friday March 3"
    m_sAppStats(4) = "Size: 77k" & vbNewLine & "Guest Speaker: Carl Sundren" & vbNewLine & "Topic: Marketing"
    m_sAppData(4) = "Notes: Competing in the new age, breaking into the new media, creating an online marketing strategy in a wired world.."

End Sub

Public Sub CreateMusicData()
    
    ReDim m_sLyrics(9)
    ReDim m_sTitles(9)
    ReDim m_sMedia(9)
    ReDim m_sDesc(9)
    
    m_sTitles(0) = "Kahmir" & vbNewLine & "Led Zeppelin"
    m_sMedia(0) = "USB Port 2, Dev ID Un1703"
    m_sDesc(0) = "Album: Physical Graffiti" & vbNewLine & "Duration: 4:10" & vbNewLine & "Accessed: 6/15/2006"
    m_sLyrics(0) = "Oh, let the sun beat down upon my face, and stars fill my dreams. I'm a traveller of both time and space, to be where I have been. To sit with elders of the gentle race, this world has seldom seen. They talk of days for which they sit and wait, all will be revealed"
    
    m_sTitles(1) = "Texas Flood" & vbNewLine & "Stevie Ray Vaughan"
    m_sMedia(1) = "CD/RW D:, GCE-451"
    m_sDesc(1) = "Album: Texas Flood" & vbNewLine & "Duration: 3:54" & vbNewLine & "Accessed: 2/17/2006"
    m_sLyrics(1) = "Well there's floodin' down in Texas, all of the telephone lines are down. Well there's floodin' down in Texas, all of the telephone lines are down. And I've been tryin' to call my baby, Lord and I can't get a single sound."
    
    m_sTitles(2) = "Creep" & vbNewLine & "Stone Temple Pilots"
    m_sMedia(2) = "Fixed Drive C:, Maxtor 60e451"
    m_sDesc(2) = "Album: Core" & vbNewLine & "Duration: 3:12" & vbNewLine & "Accessed: 10/11/2006"
    m_sLyrics(2) = "Forward yesterday, makes me wanna stay. What they said was real, makes me wanna steal. Livin' under house, guess I'm livin', I'm a mouse. All's I gots is time, got no meaning, just a rhyme"
    
    m_sTitles(3) = "Rooster" & vbNewLine & "Alice in Chains"
    m_sMedia(3) = "DVR E:, Pioneer 110D"
    m_sDesc(3) = "Album: Dirt" & vbNewLine & "Duration: 4:03" & vbNewLine & "Accessed: 1/2/2007"
    m_sLyrics(3) = "Ain't found a way to kill me yet. Eyes burn with stinging sweat. Seems every path leads me to nowhere. Wife and kids household pet. Army green was no safe bet. The bullets scream to me from somewhere. Here they come to snuff the rooster"
    
    m_sTitles(4) = "Dead and Bloated" & vbNewLine & "Stone Temple Pilots"
    m_sMedia(4) = "Fixed Drive F:, Maxtor 8192f"
    m_sDesc(4) = "Album: Core" & vbNewLine & "Duration: 3:12" & vbNewLine & "Accessed: 12/19/2006"
    m_sLyrics(4) = "I am smellin' like the rose, that somebody gave me on my birthday deathbed. I am smellin' like the rose that somebody gave me, somebody gave me, somebody gave me on my birthday deathbed"
    
    m_sTitles(5) = "Heart Shaped Box" & vbNewLine & "Nirvana"
    m_sMedia(5) = "Network Share Y:, Server DEX7"
    m_sDesc(5) = "Album: In Utero" & vbNewLine & "Duration: 3:51" & vbNewLine & "Accessed: 2/30/2005"
    m_sLyrics(5) = "She eyes me like a pisces, when I am weak. I've been locked inside, your Heart-Shaped box for a week. I was drawn into your magnet, tar pit trap. I wish I could eat your cancer, when you turn back."
    
    m_sTitles(6) = "Back in Black" & vbNewLine & "AC/DC"
    m_sMedia(6) = "Floppy Disk A:"
    m_sDesc(6) = "Album: Back in Black" & vbNewLine & "Duration: 4:24" & vbNewLine & "Accessed: 2/14/2007"
    m_sLyrics(6) = "Back in black, I hit the sack. I've been too long, I'm glad to be back. Yes, I'm let loose, from the noose, that's kept me hanging about. I've been looking at the sky, cause it's gettin' me high, forget the hearse 'cause I never die"
    
    m_sTitles(7) = "Sympathy for the devil" & vbNewLine & "Rolling Stones"
    m_sMedia(7) = "Fixed Drive C:, Maxtor 60e451"
    m_sDesc(7) = "Album: Hot Rocks II" & vbNewLine & "Duration: 3:39" & vbNewLine & "Accessed: 3/3/2005"
    m_sLyrics(7) = "Please allow me to introduce myself, I'm a man of wealth and taste. I've been around for a long long year, stolen many man's soul and faith. I was around when Jesus Christ, had His moment of doubt and pain. Made damn sure that Pilate washed his hands, and sealed His fate"
    
    m_sTitles(8) = "Behind Blue Eyes" & vbNewLine & "the Who"
    m_sMedia(8) = "CD/RW D:, GCE-451"
    m_sDesc(8) = "Album: Who's Next" & vbNewLine & "Duration: 4:01"
    m_sLyrics(8) = "No one knows what it's like, to be the bad man, to be the sad man. Behind blue eyes. No one knows what it's like, To be hated, To be fated, To telling only lies."
    
    m_sTitles(9) = "Stuck In A Moment You Cant Get Out Of" & vbNewLine & "U2"
    m_sMedia(9) = "Fixed Drive C:, Maxtor 60e451"
    m_sDesc(9) = "Album: All That You Can't Leave Behind." & vbNewLine & "Duration: 4:33" & vbNewLine & "Accessed: 5/1/2006"
    m_sLyrics(9) = "I'm not afraid Of anything in this world. There's nothing you can throw at me, that I haven't already heard. I'm just trying to find A decent melody, a song that I can sing In my own company. I never thought you were a fool, but darling look at you."

End Sub
