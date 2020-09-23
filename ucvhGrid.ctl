VERSION 5.00
Begin VB.UserControl ucVHGrid 
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MousePointer    =   99  'Custom
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ucvhGrid.ctx":0000
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   210
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   1125
   End
End
Attribute VB_Name = "ucVHGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'-> so, who has best kung-fu?


'                 _/          _/_/_/  _/_/_/    _/_/_/  _/_/_/
'    _/      _/  _/_/_/    _/        _/    _/    _/    _/    _/
'   _/      _/  _/    _/  _/  _/_/  _/_/_/      _/    _/    _/
'    _/  _/    _/    _/  _/    _/  _/    _/    _/    _/    _/
'     _/      _/    _/    _/_/_/  _/    _/  _/_/_/  _/_/_/


'********************************************************************************************
'*  vhGrid!         vhGrid - Virtual Hybrid Grid Control 1.7                                *
'*                                                                                          *
'*  Started:        November 18, 2006                                                       *
'*  Released:       Febuary 19, 2007                                                        *
'*  Updated:        July 07, 2007                                                           *
'*  Purpose:        Ultra-Fast Virtual Grid Hybrid                                          *
'*  Functions:      (partial listing)                                                       *
'*  Revision:       1.7.0                                                                   *
'*  Compile:        Native                                                                  *
'*  Author:         John Underhill (Steppenwolfe)                                           *
'*                                                                                          *
'********************************************************************************************



'-> Initialization <-
'-> GridInit RowCount, Columns
'-> AddCell Row, Column, Text, TextAlign, IconIndex, BackColor, ForeColor, Font, Indent, RowSpanDepth
'-> FastLoad Or Draw properties
'/~ These are the three core functions involved in loading grid data.
'/~ GridInit: sizes the internal arrays, and sets the grids item count. This should be called before adding data.
'/~ AddCell: loads cell data into the GridItem class array.
'/~ FastLoad: turns on drawing once the last item (specified by the GridInit parameters), has been loaded.
'/~ Draw: manually toggles the draw switch for the grid.

'-> Theming <-
'-> ThemeManager eSkinStyle, UseSkinTheme, ThemeColor, ThemeLuminence, ColumnFontColor, ColumnFontHiliteColor, ColumnFontPressedColor, OptionFormBackColor, OptionFormOffsetColor, OptionFormForeColor, OptionFormTransparency, OptionFormGradient, UseXPColors, OptionControlColor, IncludeAdvancedEdit, IncludeColumnTip, IncludeFilter, IncludeCellTip
'/~ Most style parameters can be set through this one call. Each option has a corresponding function so that option properties
'/~ may also be set individually.
'-> eSkinStyle: the skin image group
'-> UseSkinTheme: use skin colorization
'-> ThemeColor: skin colorization base
'-> ThemeLuminence: skin colorization depth
'-> ColumnFontColor, ColumnFontHiliteColor, ColumnFontPressedColor: column text colors
'-> OptionFormBackColor, OptionFormOffsetColor, OptionFormForeColor, OptionControlColor: filter, tooltip, columntip, and advanced edit forms, base color scheme
'-> OptionFormTransparency, OptionFormGradient, UseXPColors: option window effects options
'-> IncludeAdvancedEdit, IncludeColumnTip, IncludeFilter, IncludeCellTip, IncludeTreeview: options to include in the theme change
'-> ThemeAutoXp [bool]
'-> Auto assigns skin elements based on users current theme. To be usewd in place of thememenager on xp systems.

'-> Row Spanning <-
'-> AddCell :=lSpanRowDepth
'/~ Row spanning allows customizable row heights beyond limitations of the listview class. Rows can span the visible
'/~ client area, allowing for a large amount of data to be displayed on a single row. This can be used to create header
'/~ or story cells, or to emulate variable row height.

'-> Cell Spanning <-
'-> CellSpanHorizontal Row, FirstCell, LastCell
'/~ Cells can be spanned across any variation of column lengths, allowing for caption or story cells. Cell spanning
'/~ has only a modest impact on processing, and so can be used at will.

'-> SubCells <-
'-> SubCellAddControl Row, Cell, Width, Height, ControlHandle, Position, Left, Top, UseVirtualRow, UseVirtualCell
'/~ Control Subcells contain a control window. There are 15 integrated control types you can use, but a control
'/~ subcell can also use external controls by passing in the controls handle to the subcell and setting the controls
'/~ parent property to the grid hwnd. Control subcells are meant to be used sparingly, as each one requires
'/~ a unique control instance, (so say, you want 3 controls per row, 1k rows, that's 3k controls!). I would not
'/~ recommend more then 30 controls per grid instance.
'/~ I have added the ability to switch between edit controls, (on the fly?), lending greater flexibility
'/~ with that usage, but current implementation of subcells will remain as it is.

'-> Virtual Access <-
'/* Some features are not supported in virtual mode, simply because they require the GridItem class to work.
'/* Currently not supported: row spanning, cell spanning, column filter, subcells, cell tips, and cell headers.
'/* Properties or functions that rely on an internal data for calculations are also bypassed, for example
'/* ColumnTextFitHeight and Find.
'/* Cell tips and cell headers might be added later, using an events interface to collect the data.

'-> Owner Drawn Cells <-
'-> OwnerDrawImpl = 'object'
'/* Cells can be ownerdrawn through various stages of the cell rendering process. The default draw
'/* behaviors can also be bybassed by setting the bSkipdefault param to true within the callback interface.
'/* The callback returns the row and cell numbers, cell hdc, and a reference to the rows griditem class
'/* instance. The griditem variables can be changed or used as a reference in the draw process.

'-> Hyper Mode <-
'/* This mode presents several advantages to loading a data set in the standard way. First, how it works.
'/* Most applications collect data through internal methods, then once those processes complete, they load
'/* that data into a display control. The disadvantages of this method are that you are reproducing all that
'/* data a second time when populating (say a listview) internal data structs. This is a serious waste of
'/* memory, and can be a very slow process. What hyper mode does, is allows you direct access to the grids
'/* internal storage class, you populate that directly instead of using arrays, then when data collection
'/* has completed, you simply pass the array pointer into the grid. So data is not reproduced, which means
'/* only half the memory footprint, and there is no need to populate internal structures, so the time saved
'/* in loading grid can be very dramatic. On my benchmarks, a standard listview takes almost 14 seconds
'/* to load 100 thousand items with 8 columns, compare that to 1/10th of a second using hypermode.

'*** Notes ***
'-! Disclaimer
'/~ I don't usually put serious restrictions on my source, but I can see this getting some
'/~ usage so, here is the nasty bit..
'/~ Copyright © 2006 John Underhill. All rights reserved world-wide.
'/~ This software is protected under a general GNU license. Intellectual and software
'/~ Copyrights reserved.
'/~ Terms and Conditions of Use
'/~ By using this software you agree to abide by the following conditions:
'/~ 1) John Underhill, (the author), shall accept no liability or responsibility for
'/~ the use of this software. No warranty, or guarantee of fitness, or promise of support
'/~ is either expressed or implied, and no responsibility for this software, in any way
'/~ imaginable, is assumed by the author.
'/~ 2) You may use this software in your personal projects as you like. Any commercial
'/~ product using this software, must acquire the authors expressed consent (email), before
'/~ publication. I reserve the right to refuse the use of this software in situations
'/~ where I do not think the usage, or product is appropriate.
'/~ 3) This software may not be used in any product that contains malicious code, including
'/~ spyware, malware, adware, or virii.
'/~ 4) You may freely distribute this source code where appropriate, but all notes must remain
'/~ intact, and the author should be given the proper credit.

'-* Credits/Kudos
'/~ Most of all, a thanks goes out for the inspiration drawn from Steve McMahon.
'/~ The sort routines are adaptations of Rohan Edwards non-recursive triquicksort routines.
'/~ The clsImageDrag and clsImageList are rewrites of Steve's classes, (www.vbaccelrator.com).
'/~ The subclasser is a rewrite of Paul Caton's WinSubHook with some small changes and unicode support added.
'/~ Thanks to Carles P.V. for his api treeview control, the basis of the integrated treeview.
'/~ Thanks to Keith 'Lavolpe' Fox, for his keen eye in bugchecking.
'/~ Thanks to Zhu, for his ongoing bugchecks, and help with unicode support.
'/~ The rest constitutes an original work, authored by yours truly..

'-# History
'/~ Steve McMahon had his sGrid II, and I have my little vhGrid ;o)
'/~ I started this project with rewriting sGrid II in mind, but soon realized that the methods Steve used
'/~ would not allow some of the ideas I had in store for this project, (like virtual listing, and row spanning).
'/~ So, using my HyperList as a framework, I began building this grid. There were a lot of challenging
'/~ technical bits along the way, like real-time vertically sizable headers, custom tooltips, filter menu, and
'/~ the advanced edit, (but the feature that caused me the most grief by far, was the spanning, lost a week on that)..
'/~ Several of the classes are portable, like tooltips, imagelist, and odcontrol. Odcontrol is an example
'/~ of creating controls strictly with api, they run faster, with less overhead, and have many
'/~ options accessible that are hidden in the vb com implementation.
'/~ I looked at a lot of grid controls while writing this, (sourceforge and codeguru/codeproject have some
'/~ good ones), and it seems to me, that as far as features and interface design go, sGrid II, still stood out
'/~ as one of the best. So then, how to make a better grid control then sGrid II? Well, I don't know if this is
'/~ better, but it is certainly a very powerful tool, with some unique features to it, you be the judge..

'-! Cautions
'/~ After publishing my listview, I was besieged by (hundreds!) emails asking basic deployment questions,
'/~ almost every one of these questions could have been resolved by the asker, if they had only taken the time
'/~ to read the documentation, follow the examples provided, or stepped through the code to see how it was
'/~ working. With this grid, I have expanded the documentation with clear examples, and explanations of
'/~ properties/routines. There are thousands of comments in the code, which should simplify the readers understanding
'/~ of the grids mechanics. That being the case, I will not entertain questions about the grids usage, you either
'/~ take the time to figure it out, -or don't use it-.
'/~ I also saw HyperList (v1), on a chinese website, with the guy claiming he wrote it, (looks like he might have even
'/~ been trying to sell it, har~har). Now, I don't ask for much, don't care about contest, (votes are only a tip of the hat,
'/~ as it should be..), but if you use my code, (or anyone else's), show some integrity, and give the credit due the author.
'/~ Don't ask me for features! There is enough code in this already, if you want treeview in cell, or ado connector, or
'/~ grid to turn into three-dimensional rotating cube when r-clicked (w/ talking yoda ai), whatever, do it yourself!
'/~ This is what I wanted in a grid. If you need something else, add it, you may even learn something..
'/~ ..and if you don't like my attitude, then why are you reading this?

'-! Compiling and Distribution
'/~ You compile this control, just like any other usercontrol. Highlight the uc in the project explorer, then go to
'/~ Project-> vhGrid Properties and RENAME THE CONTROL! If you do not, and someone else installs software with this grid
'/~ using the same name, but a different version number, your software will-stop-working. A good idea is to prefix it
'/~ with the software or company name. Goto File -> Make "mycompanynamegrid.ocx". Then go back to properties and change
'/~ the version compatibility from 'project compatibility' to 'binary compatibility' and compile it again.

'-? Recommendations
'/~ If you only intend to use one skin, then delete the other skin elements from the resource file. This should reduce
'/~ compile size by almost half.

'-@ 98/ME users
'/~ No unicode support for legacy operating systems. 98/ME, have only basic support using glyphs (hideous looking), so
'/~ I did not bother with it. I have not tested this on legacy system, but it should be ok, if not, send me an email with
'/~ -specific details- of the problem, (where/when/what), and if you want to work with me, I will try to fix it.

'-> Using vhGrid
'/~ Yes I know it looks complicated, and there is a learning curve, but if you study examples, and take the time to read
'/~ the comments, it is not that bad. Level of complexity is related to sophistication required of display. If you want a fancy
'/~ MP3 interface with pic frame and integrated controls, od cells, etc, then you will have to do some experimenting with
'/~ properties and methods, (just like every other control I have ever used). Dig in, put a button on a form and test
'/~ the properties one by one, step through routines, study the examples, what you create is limited only by your willingness
'/~ to learn, and depth of your imagination.

'-@ Bugfixes to Ver. 1.0, Feb 20/07
'/~ Compiler hanging when building demo exe. An error during uc termination causing a jump before all objects could be
'/~ resolved, (including manifest shell reference). Tested for the error with array check in DestroyFont routine.
'/~ Text only focus on cells with cellheaders was misaligned. Rewrote section and added CellCalcHeaderSize.
'/~ Pre-population focus and click events caused jump out of grid wndproc, added rowcount checks to called routines.
'/~ Checkbox hittest failing when grid is not using icons. CellDrawIcon was resolving a clip region call, added an additional
'/~ SelectClipRgn call to clear the clip region when icons are absent.
'/~ Removed many of the default handlers used for debugs during build up, and added various conditionals to bypass errors.
'/~ Header image 'blackout' on first column resize. This was a strange one.. The cause was that a getclientrect call in ColumnRender
'/~ was returning a very large rect (32k right) the first time column was sized, CreateCompatibleBitmap call failed
'/~ creating an empty dc. Looks like some max scroll width applied to header class. Compensated by building the rect
'/~ manually, using bmp max size of 5x the grid client width.
'/~ Skinned scrollbar dissappearing after toggling visible property (compiled), added a scrollbar.refresh to usercontrol_show.
'/~ Standard cursor showing when headersizable turned off, and cursor over that hittest region, adjusted logic in header proc.
'/~ Repeated sorts with direction changes led to gridcell corruption. Caused by a an error in the sortcontrol routine,
'/~ adjusted routine logic to compensate.
'/~ Column add/remove failing when multiple columns changed. Added a resize array call into griditem class to adjust internal
'/~ arrays for active cell count changes.
'/~ Cell colors lightening with each row added. Test for color dimensioned array before applying xp offsets in CellColor sub.

'-@ Bugfixes/Additions to Ver. 1.1, Feb 23/07
'/~ Found source of header size glitch. Was in grid proc HDM_LAYOUT sizing hack. Header size was using getwindowrect
'/~ and not subtracting left coord, so header was growing with every resize. Changed to getclientrect call.
'/~ FontHandle function was using isnt switch instead of isunicode on creating logfont structure, so fonts remained as
'/~ unicode default arial. Added more unicode support to peripheral classes.
'/~ Went over the property list and sorted out the propbag settings.
'/~ Added combo/imagecombo and listbox/imagelistbox to edit controls. Edit controls can now be swapped on the
'/~ fly using the eHEditRequest event, demonstrated in the subcell demo.
'/~ Added auto-scrolling to row drag and drop functionality.
'/~ Added the ForeColor auto property, (thanks to Keith 'Lavolpe' Fox, for the code snippet).
'/~ Fixed the headerhide property by bypassing the HDM_LAYOUT message.
'/~ Rewrote portions of odcontrol class, adding blended backcolor to list and textbox, and changes to combo
'/~ codes. Moved the edit textbox out of the uc, and to an odcontrol instance.
'/~ Added right click event.

'-@ Bugfixes/Additions to Ver. 1.2, Feb 26/07
'/~ Added esc, tab, and enter accelerators to close edit window. Added spacebar accelerator to toggle checkbox.
'/~ Fixed partial cell focus when editor is unloaded.
'/~ Fixed header drop icon in drag image by hiding icon while dropping, (was showing a black mask per m$
'/~ internal imagelist issue).
'/~ Fixed incorrect row count after filter is applied.
'/~ Fixed row drop scrolling with smoother downward descent.
'/~ Now hiding sort icon and bypassing filter when row count is zero.
'/~ Added focus forecolor property to items, thememanager and filter.
'/~ Adjusted row and cell counts in gridinit routine to reflect actual numbers.
'/~ Adjusted add/remove/clear row routines for starting row index now can add/remove by row.
'/~ Enhanced the glyphs on the scrollbar buttons.
'/~ Adjusted the enabled property and added disabled backcolor/forecolor properties.
'/~ Excluded spanned rows from column filters, you can not filter spanned rows. This is for technical reasons
'/~ but also because spanned rows are not meant for standard cell data, but as story or control cells.
'/~ Tool Tiptimer stopped when grid out of focus.
'/~ Rebuilt header proc, reduced hit testing from four routines to one. Cursor changes during header
'/~ changes are fixed. Dragging and passing mouse past header caused header to start sizing, fixed with
'/~ hittesting routine.
'/~ Rewrote portions of grid proc, dropping header into list client caused misfire of timer, and dropping
'/~ rows out of client area also caused issues, both problems resolved.
'/~ Added RowNoEdit sub, selected row will not be editable.

'-@ Bugfixes/Additions to Ver. 1.3, Mar 04/07
'/~ Added integrated treeview control! Now has an optional parallel treeview with all the trimmings.
'/~ Treeview is skinned, with custom checkboxes, and can be used standalone as class or uc.
'/~ Added left align scrollbar property to treeview and grid.
'/~ clsSkinScrollbars was completely rebuilt, it is faster, leaner, and can now be transported and
'/~ used on (some) other controls, as demonstrated with treeview.
'/~ Added refresh to filter unload, and backcolor change.
'/~ Added resize event to uc in response to display or settings changes.
'/~ Fixed horz scrollbar button size bug. When hz button was less then 50 px, it dissapeared when pressed.
'/~ Added a refresh to transition mask while header is in sizing state. Header was erasing portions of mask.
'/~ Fixed a logic bug in reloading rows, was taking a long time on secondary builds. Can now reload 1k rows
'/~ in 0.1 seconds.
'/~ Added demonstration of HyperMode. It is now possible to build data arrays directly into a clsGridItem
'/~ array, and pass only the pointer into grid. This means you can use the griditems for data storage in
'/~ your application, then load results near instantaneously into grid.
'/~ Benched on athalon 1.4, 512 ram: 100k rows, 8 columns of mixed data loaded in .09 s.

'-@ Bugfixes/Additions to Ver. 1.4, Mar 18/07
'/~ Added integrated 32 bit icon support with property page. clsImagelist is portable, and with property sheet
'/~ and a couple of simple changes can add an integrated 32b imagelist to any usercontrol.
'/~ Added comlete mousewheel support.
'/~ Fixed removal of last spanned row.
'/~ Fixed a number of focus issues, including set/kill full row select, added icon focus state, switched imls
'/~ to integrated class versions, and established inter-focus events between treeview and grid.
'/~ Added right leading font support throughout.
'/~ Added vertical scrollbar alignment properties to grid and treeview.
'/~ Fixed an issue with the timer subclassing, where it was unloading message when not in table.
'/~ Treeview can now be aligned to any coordinate (N/S/W/E).
'/~ Added a manifest to demonstrate icon alphablending with compiled project (subcell demo).
'/~ Fixed problems caused by manifest in odcontrol, mostly xp rendering bypassing cntl color subclassing.
'/~ Fixed a bug in row drag & drop where drag image did not load on compiled project, by altering image capture method.
'/~ Added drag capability between treeview and grid. Tree nodes can now be dropped into grid cells.

'-@ Bugfixes/Additions to Ver. 1.5, Mar 27/07
'/~ Complete rewrite of row spanning. Now vertical spanning has almost no cpu impact and can be used
'/~ on any number of rows, with no impact on performance. Also contributing to faster load/render times
'/~ when grid is using the spanning feature.
'/~ Added pageup/down home/end to scroll methods in clsSkinScrollbars.
'/~ Added 'S' and 'E' key accelerators  that scroll to top and bottom of list, and now consuming other char input.
'/~ Moved checkbox state to griditem class, now checkboxes can be sorted.
'/~ Fixed an issue with column locking.
'/~ Finished unicode support for treeview, should now be fully compliant.
'/~ Fixed issue with cell spanning and removing columns by adjusting griditem cellspan internal counters.
'/~ Added stretchblt mode options to clsRender class.
'/~ Added built in alpha icon support, now alpha icons are available on any system with gdiplus installed.
'/~ Added three new visual styles, xp blue, xp green, and vista.
'/~ Added autoxp switch, which on xp system assigns skin elements based on current theme.
'/~ Barring any future bugfixes, this will be the last version in vb6.

'-@ Bugfixes/Additions to Ver. 1.6, July 07/07
'/~ Added four new styles, (vista/quicksilver/xp-blue/xp-green)
'/~ Added row tags: RowTag(row) = ""
'/~ Added auto xp style property: ThemeAutoXp = True
'/~ Fixed focus text only off coordinates when cell header is present.
'/~ Fixed text clipping in cells
'/~ Bunch of other stuff I can't remember..


'-> enjoy

'-@ steppenwolfe_2000@yahoo.com
'-> Cheers,
'-> John


Implements GXISubclass


Private Const BM_TRANSPARENT                    As Long = 1

Private Const CCM_FIRST                         As Long = &H2000
Private Const CCM_SETUNICODEFORMAT              As Long = (CCM_FIRST + 5)
Private Const CCM_GETUNICODEFORMAT              As Long = (CCM_FIRST + 6)

Private Const CLR_NONE                          As Long = -1

 '/* edit control styles
Private Const ES_UPPERCASE                      As Long = &H8
Private Const ES_LOWERCASE                      As Long = &H10

Private Const EM_LIMITTEXT                      As Long = &HC5

Private Const FW_NORMAL                         As Long = 400
Private Const FW_BOLD                           As Long = 700

Private Const GW_HWNDNEXT                       As Long = &H2

Private Const GWL_STYLE                         As Long = (-16)
Private Const GWL_EXSTYLE                       As Long = (-20)

Private Const H_MAX                             As Long = &HFFFF + 1

Private Const HDF_LEFT                          As Long = 0
Private Const HDF_RIGHT                         As Long = 1
Private Const HDF_CENTER                        As Long = 2
Private Const HDF_IMAGE                         As Long = &H800
Private Const HDF_BITMAP_ON_RIGHT               As Long = &H1000
Private Const HDF_STRING                        As Long = &H4000

Private Const HDI_WIDTH                         As Long = &H1
Private Const HDI_TEXT                          As Long = &H2
Private Const HDI_FORMAT                        As Long = &H4
Private Const HDI_IMAGE                         As Long = &H20

Private Const HDM_FIRST                         As Long = &H1200
Private Const HDM_GETITEMCOUNT                  As Long = (HDM_FIRST + 0)
Private Const HDM_GETITEMA                      As Long = (HDM_FIRST + 3)
Private Const HDM_SETITEMA                      As Long = (HDM_FIRST + 4)
Private Const HDM_LAYOUT                        As Long = (HDM_FIRST + 5)
Private Const HDM_GETITEMRECT                   As Long = (HDM_FIRST + 7)
Private Const HDM_SETIMAGELIST                  As Long = (HDM_FIRST + 8)
Private Const HDM_GETITEMW                      As Long = (HDM_FIRST + 11)
Private Const HDM_SETITEMW                      As Long = (HDM_FIRST + 12)
Private Const HDM_SETHOTDIVIDER                 As Long = (HDM_FIRST + 19)

Private Const HDN_FIRST                         As Long = H_MAX - 300
Private Const HDN_ITEMCHANGINGA                 As Long = (HDN_FIRST - 0)
Private Const HDN_ITEMCHANGEDA                  As Long = (HDN_FIRST - 1)
Private Const HDN_ENDTRACKA                     As Long = (HDN_FIRST - 7)
Private Const HDN_BEGINDRAG                     As Long = (HDN_FIRST - 10)
Private Const HDN_ENDDRAG                       As Long = (HDN_FIRST - 11)
Private Const HDN_ITEMCHANGINGW                 As Long = (HDN_FIRST - 20)
Private Const HDN_ITEMCHANGEDW                  As Long = (HDN_FIRST - 21)
Private Const HDN_ENDTRACKW                     As Long = (HDN_FIRST - 27)

Private Const HDR_MINHEIGHT                     As Long = 24
Private Const HDR_MAXHEIGHT                     As Long = 140

Private Const ICC_LISTVIEW_CLASSES              As Long = &H1

Private Const ILC_MASK                          As Long = &H1
Private Const ILC_COLOR32                       As Long = &H20

Private Const ILD_TRANSPARENT                   As Long = &H1

Private Const LOGPIXELSY                        As Long = 90

Private Const LF_ANTIALIASED_QUALITY            As Long = 4
Private Const LF_CLEARTYPE_QUALITY              As Long = 5

Private Const LVCF_FMT                          As Long = &H1
Private Const LVCF_WIDTH                        As Long = &H2
Private Const LVCF_TEXT                         As Long = &H4
Private Const LVCF_ORDER                        As Long = &H20

Private Const LVIF_STATE                        As Long = &H8
    
Private Const LVHT_NOWHERE                      As Long = &H1
Private Const LVHT_ONITEMICON                   As Long = &H2
Private Const LVHT_ONITEMLABEL                  As Long = &H4
Private Const LVHT_ONITEMSTATEICON              As Long = &H8
Private Const LVHT_ONITEM                       As Long = (LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_ONITEMSTATEICON)

Private Const LVIR_BOUNDS                       As Long = &H0
Private Const LVIR_LABEL                        As Long = &H2

Private Const LVIS_FOCUSED                      As Long = &H1
Private Const LVIS_SELECTED                     As Long = &H2
Private Const LVIS_CUT                          As Long = &H4
    
Private Const LVM_FIRST                         As Long = &H1000
Private Const LVM_SETBKCOLOR                    As Long = (LVM_FIRST + 1)
Private Const LVM_SETIMAGELIST                  As Long = (LVM_FIRST + 3)
Private Const LVM_GETITEMCOUNT                  As Long = (LVM_FIRST + 4)
Private Const LVM_GETITEMRECT                   As Long = (LVM_FIRST + 14)
Private Const LVM_HITTEST                       As Long = (LVM_FIRST + 18)
Private Const LVM_ENSUREVISIBLE                 As Long = (LVM_FIRST + 19)
Private Const LVM_REDRAWITEMS                   As Long = (LVM_FIRST + 21)
Private Const LVM_GETCOLUMNA                    As Long = (LVM_FIRST + 25)
Private Const LVM_SETCOLUMNA                    As Long = (LVM_FIRST + 26)
Private Const LVM_INSERTCOLUMNA                 As Long = (LVM_FIRST + 27)
Private Const LVM_DELETECOLUMN                  As Long = (LVM_FIRST + 28)
Private Const LVM_GETCOLUMNWIDTH                As Long = (LVM_FIRST + 29)
Private Const LVM_SETCOLUMNWIDTH                As Long = (LVM_FIRST + 30)
Private Const LVM_GETHEADER                     As Long = (LVM_FIRST + 31)
Private Const LVM_SETTEXTCOLOR                  As Long = (LVM_FIRST + 36)
Private Const LVM_SETTEXTBKCOLOR                As Long = (LVM_FIRST + 38)
Private Const LVM_GETTOPINDEX                   As Long = (LVM_FIRST + 39)
Private Const LVM_GETCOUNTPERPAGE               As Long = (LVM_FIRST + 40)
Private Const LVM_UPDATE                        As Long = (LVM_FIRST + 42)
Private Const LVM_SETITEMSTATE                  As Long = (LVM_FIRST + 43)
Private Const LVM_GETITEMSTATE                  As Long = (LVM_FIRST + 44)
Private Const LVM_SETITEMCOUNT                  As Long = (LVM_FIRST + 47)
Private Const LVM_GETSELECTEDCOUNT              As Long = (LVM_FIRST + 50)
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE      As Long = (LVM_FIRST + 54)
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE      As Long = (LVM_FIRST + 55)
Private Const LVM_GETSUBITEMRECT                As Long = (LVM_FIRST + 56)
Private Const LVM_SUBITEMHITTEST                As Long = (LVM_FIRST + 57)
Private Const LVM_GETCOLUMNW                    As Long = (LVM_FIRST + 95)
Private Const LVM_SETCOLUMNW                    As Long = (LVM_FIRST + 96)
Private Const LVM_INSERTCOLUMNW                 As Long = (LVM_FIRST + 97)

Private Const LVN_FIRST                         As Long = -100&
Private Const LVN_COLUMNCLICK                   As Long = (LVN_FIRST - 8)
Private Const LVN_BEGINDRAG                     As Long = (LVN_FIRST - 9)
Private Const LVN_BEGINRDRAG                    As Long = (LVN_FIRST - 11)
Private Const LVN_ENDDRAG                       As Long = (LVN_FIRST - 12) '/* undocumented?
    
Private Const LVS_REPORT                        As Long = &H1
Private Const LVS_SINGLESEL                     As Long = &H4
Private Const LVS_SHOWSELALWAYS                 As Long = &H8
Private Const LVS_SORTASCENDING                 As Long = &H10
Private Const LVS_SHAREIMAGELISTS               As Long = &H40
Private Const LVS_OWNERDRAWFIXED                As Long = &H400
Private Const LVS_OWNERDATA                     As Long = &H1000
Private Const LVS_NOCOLUMNHEADER                As Long = &H4000

Private Const LVS_EX_CHECKBOXES                 As Long = &H4&
Private Const LVS_EX_HEADERDRAGDROP             As Long = &H10&
Private Const LVS_EX_FULLROWSELECT              As Long = &H20&

Private Const LVSCW_AUTOSIZE                    As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER          As Long = -2

Private Const LVSICF_NOINVALIDATEALL            As Long = &H1

Private Const LVSIL_SMALL                       As Long = 1
Private Const LVSIL_STATE                       As Long = 2

Private Const NM_FIRST                          As Long = H_MAX
Private Const NM_CLICK                          As Long = (NM_FIRST - 2)
Private Const NM_DBLCLK                         As Long = (NM_FIRST - 3)
Private Const NM_RETURN                         As Long = (NM_FIRST - 4)
Private Const NM_RCLICK                         As Long = (NM_FIRST - 5)
Private Const NM_RDBLCLK                        As Long = (NM_FIRST - 6)
Private Const NM_SETFOCUS                       As Long = (NM_FIRST - 7)
Private Const NM_KILLFOCUS                      As Long = (NM_FIRST - 8)
Private Const NM_CUSTOMDRAW                     As Long = (NM_FIRST - 12)
Private Const NM_HOVER                          As Long = (NM_FIRST - 13)
Private Const NM_NCHITTEST                      As Long = (NM_FIRST - 14)
Private Const NM_KEYDOWN                        As Long = (NM_FIRST - 15)
Private Const NM_RELEASEDCAPTURE                As Long = (NM_FIRST - 16)
Private Const NM_SETCURSOR                      As Long = (NM_FIRST - 17)
Private Const NM_CHAR                           As Long = (NM_FIRST - 18)

Private Const PRP_APT                           As Long = 130
Private Const PRP_BRDSTL                        As Long = 1

Private Const SB_LINEDOWN                       As Long = 1
Private Const SB_LINELEFT                       As Long = 0
Private Const SB_LINERIGHT                      As Long = 1
Private Const SB_LINEUP                         As Long = 0
Private Const SB_VERT                           As Long = 1

Private Const SIF_RANGE                         As Long = &H1
Private Const SIF_PAGE                          As Long = &H2
Private Const SIF_POS                           As Long = &H4
Private Const SIF_DISABLENOSCROLL               As Long = &H8
Private Const SIF_TRACKPOS                      As Long = &H10
Private Const SIF_ALL                           As Long = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

Private Const SHGFI_ICON                        As Long = &H100
Private Const SHGFI_SYSICONINDEX                As Long = &H4000
Private Const SHGFI_LARGEICON                   As Long = &H0
Private Const SHGFI_SMALLICON                   As Long = &H1
Private Const SHGFI_OPENICON                    As Long = &H2
Private Const SHGFI_SHELLICONSIZE               As Long = &H4
Private Const SHGFI_USEFILEATTRIBUTES           As Long = &H10

Private Const SW_HIDE                           As Long = &H0
Private Const SW_NORMAL                         As Long = &H1

Private Const SWP_NOSIZE                        As Long = &H1
Private Const SWP_NOMOVE                        As Long = &H2
Private Const SWP_NOZORDER                      As Long = &H4
Private Const SWP_NOACTIVATE                    As Long = &H10
Private Const SWP_FRAMECHANGED                  As Long = &H20
Private Const SWP_SHOWWINDOW                    As Long = &H40
Private Const SWP_HIDEWINDOW                    As Long = &H80
Private Const SWP_NOOWNERZORDER                 As Long = &H200

Private Const TVGN_ROOT                         As Long = &H0
Private Const TVGN_NEXT                         As Long = &H1
Private Const TVGN_PREVIOUS                     As Long = &H2
Private Const TVGN_PARENT                       As Long = &H3
Private Const TVGN_CHILD                        As Long = &H4
Private Const TVGN_DROPHILITE                   As Long = &H8
Private Const TVGN_CARET                        As Long = &H9
Private Const TVGN_FIRSTVISIBLE                 As Long = &H5
Private Const TVGN_NEXTVISIBLE                  As Long = &H6
Private Const TVGN_PREVIOUSVISIBLE              As Long = &H7
Private Const TVGN_LASTVISIBLE                  As Long = &HA

Private Const TV_FIRST                          As Long = &H1100
Private Const TVM_INSERTITEM                    As Long = (TV_FIRST + 0)
Private Const TVM_DELETEITEM                    As Long = (TV_FIRST + 1)
Private Const TVM_EXPAND                        As Long = (TV_FIRST + 2)
Private Const TVM_GETITEMRECT                   As Long = (TV_FIRST + 4)
Private Const TVM_GETNEXTITEM                   As Long = (TV_FIRST + 10)
Private Const TVM_SELECTITEM                    As Long = (TV_FIRST + 11)
Private Const TVM_GETITEM                       As Long = (TV_FIRST + 12)
Private Const TVM_SETITEM                       As Long = (TV_FIRST + 13)
Private Const TVM_EDITLABEL                     As Long = (TV_FIRST + 14)
Private Const TVM_HITTEST                       As Long = (TV_FIRST + 17)
Private Const TVM_CREATEDRAGIMAGE               As Long = (TV_FIRST + 18)
Private Const TVM_SORTCHILDREN                  As Long = (TV_FIRST + 19)
Private Const TVM_ENSUREVISIBLE                 As Long = (TV_FIRST + 20)
Private Const TVM_ENDEDITLABELNOW               As Long = (TV_FIRST + 22)
Private Const TVM_SETINSERTMARK                 As Long = (TV_FIRST + 26)
Private Const TVM_SETINSERTMARKCOLOR            As Long = (TV_FIRST + 37)
Private Const TVM_GETINSERTMARKCOLOR            As Long = (TV_FIRST + 38)

Private Const TVN_FIRST                         As Long = -400
Private Const TVN_SELCHANGING                   As Long = (TVN_FIRST - 1)
Private Const TVN_SELCHANGED                    As Long = (TVN_FIRST - 2)
Private Const TVN_ITEMEXPANDING                 As Long = (TVN_FIRST - 5)
Private Const TVN_ITEMEXPANDED                  As Long = (TVN_FIRST - 6)
Private Const TVN_BEGINDRAGA                    As Long = (TVN_FIRST - 7)
Private Const TVN_DELETEITEM                    As Long = (TVN_FIRST - 9)
Private Const TVN_BEGINLABELEDIT                As Long = (TVN_FIRST - 10)
Private Const TVN_ENDLABELEDIT                  As Long = (TVN_FIRST - 11)
Private Const TVN_BEGINDRAGW                    As Long = (TVN_FIRST - 56)

Private Const VER_PLATFORM_WIN32_NT             As Long = 2

Private Const VK_LBUTTON                        As Long = &H1
Private Const VK_RBUTTON                        As Long = &H2
Private Const VK_TAB                            As Long = &H9
Private Const VK_ENTER                          As Long = &HD
Private Const VK_CONTROL                        As Long = &H11
Private Const VK_LEFT                           As Long = &H25
Private Const VK_RIGHT                          As Long = &H27
Private Const VK_ESCAPE                         As Long = &H1B
Private Const VK_SPACE                          As Long = &H20
Private Const VK_UP                             As Long = &H26
Private Const VK_DOWN                           As Long = &H28
Private Const VK_UCASEA                         As Long = &H41

Private Const WC_LISTVIEW                       As String = "SysListView32"

Private Const WM_SETFOCUS                       As Long = &H7
Private Const WM_KILLFOCUS                      As Long = &H8
Private Const WM_SETFONT                        As Long = &H30
Private Const WM_WINDOWPOSCHANGED               As Long = &H47
Private Const WM_PAINT                          As Long = &HF
Private Const WM_KEYDOWN                        As Long = &H100
Private Const WM_MOUSEMOVE                      As Long = &H200
Private Const WM_TIMER                          As Long = &H113&
Private Const WM_VSCROLL                        As Long = &H115
Private Const WM_HSCROLL                        As Long = &H114

Private Const WS_TABSTOP                        As Long = &H10000
Private Const WS_THICKFRAME                     As Long = &H40000
Private Const WS_HSCROLL                        As Long = &H100000
Private Const WS_VSCROLL                        As Long = &H200000
Private Const WS_BORDER                         As Long = &H800000
Private Const WS_CLIPCHILDREN                   As Long = &H2000000
Private Const WS_CLIPSIBLINGS                   As Long = &H4000000
Private Const WS_VISIBLE                        As Long = &H10000000
Private Const WS_CHILD                          As Long = &H40000000

Private Const WS_EX_WINDOWEDGE                  As Long = &H100
Private Const WS_EX_CLIENTEDGE                  As Long = &H200
Private Const WS_EX_RTLREADING                  As Long = &H2000
Private Const WS_EX_LEFTSCROLLBAR               As Long = &H4000
Private Const WS_EX_STATICEDGE                  As Long = &H20000

Public Enum EXTXpTheme
    None = 0&
    HomeStead = 1&
    NormalColor = 2&
    Metallic = 3&
End Enum

Private Enum EKNKeyNavigate
    EKNLeft = 1&
    EKNRight = 2&
    EKNUp = 3&
    EKNDown = 4&
End Enum

Private Enum TT_NOTIFICATIONS
    TTN_FIRST = -520&
    TTN_LAST = -549&
    TTN_GETDISPINFO = (TTN_FIRST - 0)
End Enum

Private Enum SYSTEM_METRICS
    SM_CXSCREEN = 0&
    SM_CYSCREEN = 1&
    SM_CXVSCROLL = 2&
    SM_CYHSCROLL = 3&
    SM_CYCAPTION = 4&
    SM_CXBORDER = 5&
    SM_CYBORDER = 6&
    SM_CYVTHUMB = 9&
    SM_CXHTHUMB = 10&
    SM_CXICON = 11&
    SM_CYICON = 12&
    SM_CXCURSOR = 13&
    SM_CYCURSOR = 14&
    SM_CYMENU = 15&
    SM_CXFULLSCREEN = 16&
    SM_CYFULLSCREEN = 17&
    SM_CYKANJIWINDOW = 18&
    SM_MOUSEPRESENT = 19&
    SM_CYVSCROLL = 20&
    SM_CXHSCROLL = 21&
    SM_CXMIN = 28&
    SM_CYMIN = 29&
    SM_CXSIZE = 30&
    SM_CYSIZE = 31&
    SM_CXFRAME = 32&
    SM_CYFRAME = 33&
    SM_CXMINTRACK = 34&
    SM_CYMINTRACK = 35&
    SM_CXSMICON = 49&
    SM_CYSMICON = 50&
    SM_CYSMCAPTION = 51&
    SM_CXMINIMIZED = 57&
    SM_CYMINIMIZED = 58&
    SM_CXMAXTRACK = 59&
    SM_CYMAXTRACK = 60&
    SM_CXMAXIMIZED = 61&
    SM_CYMAXIMIZED = 62&
End Enum

Public Enum ECTTextAlignFlags
    DT_TOP = &H0&
    DT_LEFT = &H0&
    DT_CENTER = &H1&
    DT_RIGHT = &H2&
    DT_VCENTER = &H4&
    DT_BOTTOM = &H8&
    DT_WORDBREAK = &H10&
    DT_SINGLELINE = &H20&
    DT_EXPANDTABS = &H40&
    DT_TABSTOP = &H80&
    DT_NOCLIP = &H100&
    DT_EXTERNALLEADING = &H200&
    DT_CALCRECT = &H400&
    DT_NOPREFIX = &H800&
    DT_INTERNAL = &H1000&
    DT_EDITCONTROL = &H2000&
    DT_PATH_ELLIPSIS = &H4000&
    DT_END_ELLIPSIS = &H8000&
    DT_MODIFYSTRING = &H10000
    DT_RTLREADING = &H20000
    DT_WORD_ELLIPSIS = &H40000
End Enum

Public Enum ECHCellHiliteStyle
    echStripe = 0&
    echThin = 1&
    echThick = 2&
End Enum

Public Enum ETVOLEDragConstants
    tvdNone = 0&
    tvdManual = 1&
End Enum

Public Enum ETVTreeViewAlignment
    etvLeftAlign = 0&
    etvRightAlign = 1&
    etvTopAlign = 2&
    etvBottomAlign = 3&
End Enum

Public Enum ESCScrollBarAlignment
    escRightAlign = 0&
    escLeftAlign = 1&
End Enum

Public Enum ETVNodeRelation
    trnLast = 0&
    trnFirst = 1&
    trnSort = 2&
    trnNext = 3&
    trnPrevious = 4&
End Enum

Public Enum ECTEditControlType
    ectTextBox = 0&
    ectCombo = 1&
    ectImageCombo = 2&
    ectListbox = 3&
    ectImageListbox = 4&
End Enum

Public Enum eSDSortDirection
    esdDescending = -1&
    esdDefault = 0&
    esdAscending = 1&
End Enum

Public Enum EGDDrawStage
    edgPreDraw = 1&
    edgBeforeBackGround = 2&
    edgBeforeIcon = 3&
    edgBeforeText = 4&
    edgPostDraw = 5&
End Enum

Public Enum EVSFrameStyle
    evsNoBorder = 0&
    evsThinBorder = 1&
    evsInsetBorder = 2&
    evsRaisedBorder = 3&
    evsThickBorder = 4&
End Enum

Public Enum EVSFramePosition
    evsUserDefine = 0&
    evsTopLeft = 1&
    evsTopCenter = 2&
    evsTopRight = 3&
    evsCenterLeft = 4&
    evsCenterCell = 5&
    evsCenterRight = 6&
    evsBottomLeft = 7&
    evsBottomCenter = 8&
    evsBottomRight = 9&
End Enum

Public Enum EVSTextFormat
    evsLeftAlign = 0&
    evsTopAlign = 1&
    evsRightAlign = 2&
    evsBottomAlign = 3&
End Enum

Public Enum EVSControlType
    evsCheckBox = 1&
    evsComboDropDown = 2&
    evsComboDropList = 3&
    evsComboSimple = 4&
    evsCommandButton = 5&
    evsImageCombo = 6&
    evsImageListBox = 7&
    evsLabel = 8&
    evsListBox = 9&
    evsListBoxExtended = 10&
    evsListBoxMultiSelect = 11&
    evsOptionButton = 12&
    evsPictureBox = 13&
    evsTextBox = 14&
End Enum

Public Enum EVSFrameConnector
    evsTopCap = 1&
    evsJoined = 2&
    evsBottomCap = 3&
End Enum

Public Enum EVSThemeStyle
    evsAzure = 0&
    evsClassic = 1&
    evsGloss = 2&
    evsMetallic = 3&
    evsXpSilver = 4&
    evsXpBlue = 5&
    evsXpGreen = 6&
    evsVistaArrow = 7&
    evsSilver = 8&
End Enum

Public Enum ESTSortType
    estNone = -1
    estCaseSensitive = vbBinaryCompare
    estCaseInsensitive = vbTextCompare
End Enum

Public Enum EBSBorderStyle
    ebsNone = 0&
    ebsThin = 1&
    ebsThick = 2&
End Enum

Public Enum ERDCellDecoration
    erdCellLine = 0&
    erdCellSplit = 1&
    erdCellBiLinear = 2&
    erdCellChecker = 3&
End Enum

Public Enum ECUColumnAutosize
    ecuColumnItem = LVSCW_AUTOSIZE
    ecuColumnHeader = LVSCW_AUTOSIZE_USEHEADER
End Enum

Public Enum ECAColumnAlign
    ecaColumnLeft = HDF_LEFT
    ecaColumnright = HDF_RIGHT
    ecaColumnCenter = HDF_CENTER
End Enum

Public Enum ECSColumnSortTags
    ecsSortNone = -1&
    ecsSortDefault = 0&
    ecsSortDate = 1
    ecsSortNumeric = 2&
    ecsSortIcon = 3&
    ecsSortAuto = 4&
End Enum

Public Enum EDCDropConstants
    vbOLEDropNone
    vbOLEDropManual
End Enum

Public Enum EDSDragEffectStyle
    edsClientArrow = 0&
    edsThinLine = 1&
    edsThickLine = 2&
End Enum

Public Enum EGLGridLines
    EGLNone = 0&
    EGLHorizontal = 1&
    EGLVertical = 2&
    EGLBoth = 3&
End Enum

Public Enum ECPIconPosition
    epiCenter = 0&
    epiTop = 1&
    epiBottom = 2&
End Enum

Public Enum ESTThemeLuminence
    estThemeSoft = 0&
    estThemePastel = 1&
    estThemeHard = 2&
End Enum

Public Enum EHTTextEffect
    hteTextNormal = 0&
    hteTextEmbossed = 1&
    hteTextEngraved = 2&
End Enum

Public Enum ECHTrackDepth
    ehtNarrow = 0&
    ehtWide = 1&
End Enum

Public Enum ETTToolTipPosition
    etpRightBottom = 0&
    etpRightCenter = 1&
    etpRightTop = 2&
    etpLeftBottom = 3&
    etpLeftCenter = 4&
    etpLeftTop = 5&
End Enum


Private Type tagINITCOMMONCONTROLSEX
    dwSize                                      As Long
    dwICC                                       As Long
End Type

Private Type RECT
    left                                        As Long
    top                                         As Long
    Right                                       As Long
    Bottom                                      As Long
End Type

Private Type POINTAPI
    x                                           As Long
    y                                           As Long
End Type

Private Type DRAWITEMSTRUCT
    CtlType                                     As Long
    CtlID                                       As Long
    itemID                                      As Long
    itemAction                                  As Long
    itemState                                   As Long
    hwndItem                                    As Long
    hdc                                         As Long
    rcItem                                      As RECT
    itemData                                    As Long
End Type

Private Type LVHITTESTINFO
    pt                                          As POINTAPI
    flags                                       As Long
    iItem                                       As Long
    iSubItem                                    As Long
End Type

Private Type LVCOLUMN
    Mask                                        As Long
    fmt                                         As Long
    cx                                          As Long
    pszText                                     As Long
    cchTextMax                                  As Long
    iSubItem                                    As Long
    iImage                                      As Long
    iOrder                                      As Long
End Type

Private Type HDITEM
    Mask                                        As Long
    cxy                                         As Long
    pszText                                     As String
    hbm                                         As Long
    cchTextMax                                  As Long
    fmt                                         As Long
    lParam                                      As Long
    iImage                                      As Long
    iOrder                                      As Long
End Type

Private Type HDITEMW
    Mask                                        As Long
    cxy                                         As Long
    pszText                                     As Long
    hbm                                         As Long
    cchTextMax                                  As Long
    fmt                                         As Long
    lParam                                      As Long
    iImage                                      As Long
    iOrder                                      As Long
End Type

Private Type LVITEM
    Mask                                        As Long
    iItem                                       As Long
    iSubItem                                    As Long
    State                                       As Long
    stateMask                                   As Long
    pszText                                     As Long
    cchTextMax                                  As Long
    iImage                                      As Long
    lParam                                      As Long
    iIndent                                     As Long
End Type

Private Type LOGFONT
    lfHeight                                    As Long
    lfWidth                                     As Long
    lfEscapement                                As Long
    lfOrientation                               As Long
    lfWeight                                    As Long
    lfItalic                                    As Byte
    lfUnderline                                 As Byte
    lfStrikeOut                                 As Byte
    lfCharSet                                   As Byte
    lfOutPrecision                              As Byte
    lfClipPrecision                             As Byte
    lfQuality                                   As Byte
    lfPitchAndFamily                            As Byte
    lfFaceName(32)                              As Byte
End Type

Private Type OSVERSIONINFO
    dwVersionInfoSize                           As Long
    dwMajorVersion                              As Long
    dwMinorVersion                              As Long
    dwBuildNumber                               As Long
    dwPlatformId                                As Long
    szCSDVersion(0 To 127)                      As Byte
End Type

Private Type MEASUREITEMSTRUCT
    CtlType                                     As Long
    CtlID                                       As Long
    itemID                                      As Long
    itemWidth                                   As Long
    ItemHeight                                  As Long
    itemData                                    As Long
End Type

Private Type WINDOWPOS
    hwnd                                        As Long
    hWndInsertAfter                             As Long
    x                                           As Long
    y                                           As Long
    cx                                          As Long
    cy                                          As Long
    flags                                       As Long
End Type

Private Type TEXTMETRIC
    tmHeight                                    As Long
    tmAscent                                    As Long
    tmDescent                                   As Long
    tmInternalLeading                           As Long
    tmExternalLeading                           As Long
    tmAveCharWidth                              As Long
    tmMaxCharWidth                              As Long
    tmWeight                                    As Long
    tmOverhang                                  As Long
    tmDigitizedAspectX                          As Long
    tmDigitizedAspectY                          As Long
    tmFirstChar                                 As Byte
    tmLastChar                                  As Byte
    tmDefaultChar                               As Byte
    tmBreakChar                                 As Byte
    tmItalic                                    As Byte
    tmUnderlined                                As Byte
    tmStruckOut                                 As Byte
    tmPitchAndFamily                            As Byte
    tmCharSet                                   As Byte
End Type

Private Type NMHDR
    hwndFrom                                    As Long
    idfrom                                      As Long
    code                                        As Long
End Type

Private Type NMHEADER
    hdr                                         As NMHDR
    iItem                                       As Long
    iButton                                     As Long
    lPtrHDItem                                  As Long
End Type

Private Type HDLAYOUT
    lprc                                        As Long
    lpwpos                                      As Long
End Type

Private Type NMLISTVIEW
    hdr                                         As NMHDR
    iItem                                       As Long
    iSubItem                                    As Long
    uNewState                                   As Long
    uOldState                                   As Long
    uChanged                                    As Long
    ptAction                                    As POINTAPI
    lParam                                      As Long
End Type

Private Type NMLVKEYDOWN
    hdr                                         As NMHDR
    wVKey                                       As Integer
    flags1                                      As Integer
    flags2                                      As Integer
End Type

Private Type PAINTSTRUCT
    hdc                                         As Long
    fErase                                      As Boolean
    rcPaint                                     As RECT
    fRestore                                    As Boolean
    fIncUpdate                                  As Boolean
    rgbReserved(32)                             As Byte
End Type


Private Type TVITEMW
    Mask                                        As Long
    hItem                                       As Long
    State                                       As Long
    stateMask                                   As Long
    pszText                                     As Long
    cchTextMax                                  As Long
    iImage                                      As Long
    iSelectedImage                              As Long
    cChildren                                   As Long
    lParam                                      As Long
End Type

Private Type NMTREEVIEW
    hdr                                         As NMHDR
    action                                      As Long
    itemOld                                     As TVITEMW
    itemNew                                     As TVITEMW
    ptDrag                                      As POINTAPI
End Type

Private Type SCROLLINFO
    cbSize                                      As Long
    fMask                                       As Long
    nMin                                        As Long
    nMax                                        As Long
    nPage                                       As Long
    nPos                                        As Long
    nTrackPos                                   As Long
End Type


Private Declare Function SendMessageA Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal wMsg As Long, _
                                                    ByVal wParam As Long, _
                                                    lParam As Any) As Long

Private Declare Function SendMessageW Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal wMsg As Long, _
                                                    ByVal wParam As Long, _
                                                    lParam As Any) As Long

Private Declare Function SendMessageLongA Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                             ByVal wMsg As Long, _
                                                                             ByVal wParam As Long, _
                                                                             ByVal lParam As Long) As Long

Private Declare Function SendMessageLongW Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, _
                                                                             ByVal wMsg As Long, _
                                                                             ByVal wParam As Long, _
                                                                             ByVal lParam As Long) As Long

Private Declare Function PostMessageA Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal wMsg As Long, _
                                                    ByVal wParam As Long, _
                                                    ByVal lParam As Long) As Long

Private Declare Function PostMessageW Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal wMsg As Long, _
                                                    ByVal wParam As Long, _
                                                    ByVal lParam As Long) As Long

Private Declare Function CreateFontIndirectA Lib "gdi32" (lpLogFont As LOGFONT) As Long

Private Declare Function CreateFontIndirectW Lib "gdi32" (lpLogFont As LOGFONT) As Long

Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Private Declare Function lstrcpyW Lib "kernel32" (ByVal lpString1 As Long, _
                                                  ByVal lpString2 As Long) As Long

Private Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, _
                                                 ByVal lpStr As String, _
                                                 ByVal nCount As Long, _
                                                 lpRect As RECT, _
                                                 ByVal wFormat As Long) As Long

Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, _
                                                 ByVal lpStr As Long, _
                                                 ByVal nCount As Long, _
                                                 lpRect As RECT, _
                                                 ByVal wFormat As Long) As Long

Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long, _
                                                      ByVal dwNewLong As Long) As Long

Private Declare Function SetWindowLongW Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long, _
                                                      ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowTextLengthA Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function GetWindowTextLengthW Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function GetWindowTextA Lib "user32.dll" (ByVal hwnd As Long, _
                                                          ByVal lpString As String, _
                                                          ByVal cch As Long) As Long

Private Declare Function GetWindowTextW Lib "user32.dll" (ByVal hwnd As Long, _
                                                          ByVal lpString As Long, _
                                                          ByVal cch As Long) As Long

Private Declare Function CreateWindowExA Lib "user32" (ByVal dwExStyle As Long, _
                                                       ByVal lpClassName As String, _
                                                       ByVal lpWindowName As String, _
                                                       ByVal dwStyle As Long, _
                                                       ByVal x As Long, _
                                                       ByVal y As Long, _
                                                       ByVal nWidth As Long, _
                                                       ByVal nHeight As Long, _
                                                       ByVal hWndParent As Long, _
                                                       ByVal hMenu As Long, _
                                                       ByVal hInstance As Long, _
                                                       lpParam As Any) As Long

Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, _
                                                       ByVal lpClassName As Long, _
                                                       ByVal lpWindowName As Long, _
                                                       ByVal dwStyle As Long, _
                                                       ByVal x As Long, _
                                                       ByVal y As Long, _
                                                       ByVal nWidth As Long, _
                                                       ByVal nHeight As Long, _
                                                       ByVal hWndParent As Long, _
                                                       ByVal hMenu As Long, _
                                                       ByVal hInstance As Long, _
                                                       lpParam As Any) As Long

Private Declare Function GetWindowLongA Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long) As Long

Private Declare Function GetWindowLongW Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInfo As OSVERSIONINFO) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, _
                                                                     lpSrc As Any, _
                                                                     ByVal Length As Long)

Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function ImageList_Add Lib "comctl32" (ByVal hImageList As Long, _
                                                       ByVal hBitmap As Long, _
                                                       ByVal hBitmapMask As Long) As Long

Private Declare Function ImageList_AddIcon Lib "comctl32" (ByVal hImageList As Long, _
                                                           ByVal hicon As Long) As Long

Private Declare Function ImageList_AddMasked Lib "comctl32" (ByVal hImageList As Long, _
                                                             ByVal hbmImage As Long, _
                                                             ByVal crMask As Long) As Long

Private Declare Function ImageList_Create Lib "comctl32" (ByVal MinCx As Long, _
                                                          ByVal MinCy As Long, _
                                                          ByVal flags As Long, _
                                                          ByVal cInitial As Long, _
                                                          ByVal cGrow As Long) As Long

Private Declare Function ImageList_Destroy Lib "comctl32" (ByVal hImageList As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal hIml As Long, _
                                                            ByVal I As Long, _
                                                            ByVal hdcDst As Long, _
                                                            ByVal x As Long, _
                                                            ByVal y As Long, _
                                                            ByVal fStyle As Long) As Long

Private Declare Function ImageList_GetImageCount Lib "comctl32" (ByVal hImageList As Long) As Long

Private Declare Function ImageList_Remove Lib "comctl32.dll" (ByVal hIml As Long, _
                                                              ByVal I As Long) As Long

Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, _
                                                           ByVal HPALETTE As Long, _
                                                           pccolorref As Long) As Long

Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal fEnable As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal x As Long, _
                                                    ByVal y As Long, _
                                                    ByVal cx As Long, _
                                                    ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Sub CopyMemBv Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Any, _
                                                                    ByVal pSrc As Any, _
                                                                    ByVal lByteLen As Long)

Private Declare Sub CopyMemBr Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                    pSrc As Any, _
                                                                    ByVal lByteLen As Long)

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As SYSTEM_METRICS) As Long

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Function InitCommonControlsEx Lib "comctl32" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean

Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal wCmd As Long) As Long

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal nIndex As Long) As Long

Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, _
                                                ByVal nNumerator As Long, _
                                                ByVal nDenominator As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, _
                                                      lpPoint As POINTAPI) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal crColor As Long) As Long

Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, _
                                                ByVal nBkMode As Long) As Long

Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetFocus Lib "user32" () As Long

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long) As Long

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, _
                                                ByVal nIDEvent As Long, _
                                                ByVal uElapse As Long, _
                                                ByVal lpTimerFunc As Long) As Long

Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal nIDEvent As Long) As Long

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
                                                ByVal nWidth As Long, _
                                                ByVal crColor As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long, _
                                               lpPoint As POINTAPI) As Long

Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, _
                                             ByVal x As Long, _
                                             ByVal y As Long) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                  ByVal x As Long, _
                                                  ByVal y As Long) As Long

Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, _
                                                 lpRect As RECT, _
                                                 ByVal hBrush As Long) As Long

Private Declare Function EraseRect Lib "user32" Alias "InvalidateRect" (ByVal hwnd As Long, _
                                                                        lpRect As RECT, _
                                                                        ByVal bErase As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long

Private Declare Function PtInRect Lib "user32" (lpRect As RECT, _
                                                ByVal ptX As Long, _
                                                ByVal ptY As Long) As Long

Private Declare Function InflateRect Lib "user32" (lpRect As RECT, _
                                                   ByVal x As Long, _
                                                   ByVal y As Long) As Long

Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, _
                                                lpSourceRect As RECT) As Long

Private Declare Function GetTextMetricsA Lib "gdi32" (ByVal hdc As Long, _
                                                      lpMetrics As TEXTMETRIC) As Long

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, _
                                                    ByVal y1 As Long, _
                                                    ByVal x2 As Long, _
                                                    ByVal y2 As Long) As Long

Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal hRgn As Long) As Long

Private Declare Function ExcludeClipRect Lib "gdi32" (ByVal hdc As Long, _
                                                      ByVal X1 As Long, _
                                                      ByVal y1 As Long, _
                                                      ByVal x2 As Long, _
                                                      ByVal y2 As Long) As Long

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, _
                                                  ByVal nCmdShow As Long) As Long

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, _
                                                  lpPaint As PAINTSTRUCT) As Long

Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, _
                                                lpPaint As PAINTSTRUCT) As Long

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long, _
                                               ByVal crColor As Long) As Long

Private Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hwnd As Long, _
                                                           ByVal pszSubAppName As Long, _
                                                           ByVal pszSubIdList As Long) As Long

Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, _
                                                          ByVal pszClassList As Long) As Long

Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long

Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As Long, _
                                                                ByVal dwMaxNameChars As Long, _
                                                                ByVal pszColorBuff As Long, _
                                                                ByVal cchMaxColorChars As Long, _
                                                                ByVal pszSizeBuff As Long, _
                                                                ByVal cchMaxSizeChars As Long) As Long


'/* grid events
Public Event eVHColumnAdded(ByVal lColumn As Long, ByVal lWidth As Long, ByVal lIcon As Long, ByVal sText As String) '
Public Event eVHColumnRemoved(ByVal lColumn As Long)
Public Event eVHColumnClick(ByVal lColumn As Long)
Public Event eVHColumnHorizontalSize(ByVal lColumn As Long)
Public Event eVHColumnVerticalSize(ByVal lHeight As Long)
Public Event eVHColumnDragging(ByVal lColumn As Long)
Public Event eVHColumnDragComplete() '
Public Event eVHGridSizeChange(ByVal lWidth As Long, ByVal lHeight As Long)
Public Event eVHGridEnable(ByVal bState As Boolean)
Public Event eVHItemClick(ByVal lRow As Long, ByVal lCell As Long)
Public Event eVHItemDoubleClick(ByVal lRow As Long, ByVal lCell As Long)
Public Event eVHItemRightClick(ByVal lItem As Long, ByVal lX As Long, ByVal lY As Long)
Public Event eVHItemCheck(ByVal lRow As Long, ByVal bState As Boolean)
Public Event eVHItemClear()
Public Event eVHItemDeleted(ByVal lRow As Long)
Public Event eVHItemDragging(ByVal lRow As Long)
Public Event eVHItemDragComplete(ByVal lSource As Long, ByVal lTarget As Long)
Public Event eVHVirtualAccess(ByVal lRow As Long, ByVal lCell As Long, sText As String, lIcon As Long)
Public Event eVHEditRequest(ByVal lRow As Long, ByVal lCell As Long)
Public Event eVHEditorLoaded(ByVal eEditorType As ECTEditControlType, ByVal lRow As Long, ByVal lCell As Long)
Public Event eVHEditRequestText(ByVal lRow As Long, ByVal lCell As Long, sText As String)
Public Event eVHEditChange(ByVal lRow As Long, ByVal lCell As Long, ByVal sText As String)
Public Event eVHAdvancedEditRequest(ByVal lRow As Long, ByVal lCell As Long)
Public Event eVHAdvancedEditRequestText(ByVal lRow As Long, ByVal lCell As Long, lIcon As Long, sText As String)
Public Event eVHAdvancedEditChange(ByVal lRow As Long, ByVal lCell As Long, ByVal lIcon As Long, ByVal lBackColor As Long, ByVal lForeColor As Long, ByVal sText As String)
Public Event eVHErrCond(ByVal sRtn As String, ByVal lErr As Long)
'/* treeview events
Public Event eTVClick()
Public Event eTVNodeClick(ByVal hNode As Long)
Public Event eTVNodeCheck(ByVal hNode As Long)
Public Event eTVNodeDblClick(ByVal hNode As Long)
Public Event eTVSelectionChanged()
Public Event eTVBeforeExpand(ByVal hNode As Long, ByVal ExpandedOnce As Boolean)
Public Event eTVAfterExpand(ByVal hNode As Long, ByVal ExpandedOnce As Boolean)
Public Event eTVCollapse(ByVal hNode As Long)
Public Event eTVBeforeLabelEdit(ByVal hNode As Long, Cancel As Long)
Public Event eTVAfterLabelEdit(ByVal hNode As Long, Cancel As Long, NewString As String)
Public Event eTVKeyDown(KeyCode As Long, Shift As Long)
Public Event eTVKeyPress(KeyAscii As Long)
Public Event eTVKeyUp(KeyCode As Long, Shift As Long)
Public Event eTVMouseDown(Button As Long, Shift As Long, x As Long, y As Long)
Public Event eTVMouseMove(Button As Long, Shift As Long, x As Long, y As Long)
Public Event eTVMouseUp(Button As Long, Shift As Long, x As Long, y As Long)
Public Event eTVMouseEnter()
Public Event eTVMouseLeave()

Private m_bteAlphaTransparency                  As Byte
Private m_sngLuminence                          As Single
Private m_bAdvancedEdit                         As Boolean
Private m_bEditorLoaded                         As Boolean
Private m_bAlphaBarTheme                        As Boolean
Private m_bAlphaBlend                           As Boolean
Private m_bAlphaIsLoaded                        As Boolean
Private m_bAlphaSelectorBar                     As Boolean
Private m_bCellDecoration                       As Boolean
Private m_bCellEdit                             As Boolean
Private m_bCellHotTrack                         As Boolean
Private m_bCellTips                             As Boolean
Private m_bCellTipGradient                      As Boolean
Private m_bCellTipMultiline                     As Boolean
Private m_bCellTipXPColors                      As Boolean
Private m_bCellUseXP                            As Boolean
Private m_bCheckBoxes                           As Boolean
Private m_bColumnDragging                       As Boolean
Private m_bColumnDragLine                       As Boolean
Private m_bColumnFilters                        As Boolean
Private m_bColumnFocus                          As Boolean
Private m_bColumnsShuffled                      As Boolean
Private m_bColumnSizingVertical                 As Boolean
Private m_bColumnSizingHorizontal               As Boolean
Private m_bColumnVerticalText                   As Boolean
Private m_bCustomCursors                        As Boolean
Private m_bDragDrop                             As Boolean
Private m_bDraw                                 As Boolean
Private m_bDoubleBuffer                         As Boolean
Private m_bEnabled                              As Boolean
Private m_bEditBlendBackground                  As Boolean
Private m_bFiltered                             As Boolean
Private m_bFilterLoaded                         As Boolean
Private m_bFilterGradient                       As Boolean
Private m_bFilterXPColors                       As Boolean
Private m_bFirstRowReserved                     As Boolean
Private m_bFastLoad                             As Boolean
Private m_bFullRowSelect                        As Boolean
Private m_bFocusTextOnly                        As Boolean
Private m_bGridFocused                          As Boolean
Private m_bHasSubCells                          As Boolean
Private m_bHeaderFixed                          As Boolean
Private m_bHeaderFlat                           As Boolean
Private m_bHeaderHide                           As Boolean
Private m_bHasInitialized                       As Boolean
Private m_bHeaderSizable                        As Boolean
Private m_bHeaderSizing                         As Boolean
Private m_bItemActive                           As Boolean
Private m_bIsNt                                 As Boolean
Private m_bIsXp                                 As Boolean
Private m_bLockFirstColumn                      As Boolean
Private m_bRowDragging                          As Boolean
Private m_bSorted                               As Boolean
Private m_bSkinHeader                           As Boolean
Private m_bSkinScrollBars                       As Boolean
Private m_bShowing                              As Boolean
Private m_bStopSearch                           As Boolean
Private m_bTimerActive                          As Boolean
Private m_bTipTimerActive                       As Boolean
Private m_bTipTracking                          As Boolean
Private m_bTransitionMask                       As Boolean
Private m_bUseSorted                            As Boolean
Private m_bUseThemeColors                       As Boolean
Private m_bUseCheckBoxTheme                     As Boolean
Private m_bUseUnicode                           As Boolean
Private m_bVirtualMode                          As Boolean
Private m_bVisible                              As Boolean
Private m_bXPColors                             As Boolean
Private m_bForeColorAuto                        As Boolean
Private m_bPainting                             As Boolean
Private m_bHasTreeView                          As Boolean
Private m_bTreeViewSizeable                     As Boolean
Private m_bTreeViewSizing                       As Boolean
Private m_bFontRightLeading                     As Boolean
Private m_bNodeDragging                         As Boolean
Private m_bUseSpannedRows                       As Boolean
Private m_bThemeAutoXp                          As Boolean
Private m_bImlUseAlphaIcons                     As Boolean
Private m_bIconNoHilite                         As Boolean
Private m_bFilterInvert                         As Boolean
Private m_bFilterHeader                         As Boolean
Private m_bFilterHideExact                      As Boolean
Private m_lHeaderOffset                         As Long
Private m_lTVHwnd                               As Long
Private m_lhNodeDrag                            As Long
Private m_lInitialHeaderHeight                  As Long
Private m_lDraggedColumn                        As Long
Private m_lLastDropIdx                          As Long
Private m_lStoredIcon                           As Long
Private m_lStoreHeaderHeight                    As Long
Private m_lAdvancedEditThemeColor               As Long
Private m_lAdvancedEditOffsetColor              As Long
Private m_lCellTipDelayTime                     As Long
Private m_lCellTipPosition                      As Long
Private m_lCellTipTransparency                  As Long
Private m_lCellTipVisibleTime                   As Long
Private m_lCellFocused                          As Long
Private m_lCellDepth                            As Long
Private m_lCellColorBase                        As Long
Private m_lCellColorOffset                      As Long
Private m_lIconPosition                         As Long
Private m_lColumnCount                          As Long
Private m_lColumnDivider                        As Long
Private m_lColumnSelected                       As Long
Private m_lColumnFilter                         As Long
Private m_lColumnHeight                         As Long
Private m_lCheckHeight                          As Long
Private m_lCheckWidth                           As Long
Private m_lDragImgIml                           As Long
Private m_lEditItem                             As Long
Private m_lEditHwnd                             As Long
Private m_lEditSubItem                          As Long
Private m_lFont                                 As Long
Private m_lFilterTransparency                   As Long
Private m_lHdrHwnd                              As Long
Private m_lhMod                                 As Long
Private m_lHeaderHeight                         As Long
Private m_lHGHwnd                               As Long
Private m_lHeaderTextEffect                     As Long
Private m_lImlHdHndl                            As Long
Private m_lImlRowHndl                           As Long
Private m_lImlStateHndl                         As Long
Private m_lLastCell                             As Long
Private m_lLastX                                As Long
Private m_lLastY                                As Long
Private m_lLastRow                              As Long
Private m_lParentHwnd                           As Long
Private m_lPtrOwnerDraw                         As Long
Private m_lRowMinHgt                            As Long
Private m_lRowHeight                            As Long
Private m_lRowFocused                           As Long
Private m_lRowDragEffect                        As Long
Private m_lRowIconX                             As Long
Private m_lRowIconY                             As Long
Private m_lRowCount                             As Long
Private m_lSafeTimer                            As Long
Private m_lStrctPtr                             As Long
Private m_lDisabledBackColor                    As Long
Private m_lDisabledForeColor                    As Long
Private m_lTipTimer                             As Long
Private m_lCellHiliteColor                      As Long
Private m_bColumnLock()                         As Boolean
Private m_lFilter()                             As Long
Private m_lSortArray()                          As Long
Private m_lCellColor()                          As Long
Private m_hFontHnd()                            As Long
Private m_lRowDepth()                           As Long
Private m_sSortArray()                          As String
Private m_sFilterItem()                         As String
Private m_eCellHiliteStyle                      As ECHCellHiliteStyle
Private m_eTvControlAlignment                   As ETVTreeViewAlignment
Private m_eTvScrollBarAlignment                 As ESCScrollBarAlignment
Private m_eScrollBarAlignment                   As ESCScrollBarAlignment
Private m_eEditControlType                      As ECTEditControlType
Private m_eEditFrameStyle                       As EBSBorderStyle
Private m_eHeaderSkinStyle                      As EVSThemeStyle
Private m_eThemeLuminence                       As ESTThemeLuminence
Private m_eScrollBarSkinStyle                   As EVSThemeStyle
Private m_eBorderStyle                          As EBSBorderStyle
Private m_eSortTag                              As ECSColumnSortTags
Private m_eCheckBoxSkinStyle                    As EVSThemeStyle
Private m_eGridLines                            As EGLGridLines
Private m_eHotTrackDepth                        As ECHTrackDepth
Private m_eCellDecoration                       As ERDCellDecoration
Private m_eDragEffectStyle                      As EDSDragEffectStyle
Private m_eSortType                             As ESTSortType
Private m_eOLEDragMode                          As OLEDragConstants
Private m_eEditorThemeStyle                     As EVSThemeStyle
Private m_oFilterTitleColor                     As OLE_COLOR
Private m_oCellTipForeColor                     As OLE_COLOR
Private m_oCellTipColor                         As OLE_COLOR
Private m_oCellTipOffsetColor                   As OLE_COLOR
Private m_oColumnFocusColor                     As OLE_COLOR
Private m_oBackColor                            As OLE_COLOR
Private m_oForeColor                            As OLE_COLOR
Private m_oHdrForeClr                           As OLE_COLOR
Private m_oHdrHighLiteClr                       As OLE_COLOR
Private m_oHdrPressedClr                        As OLE_COLOR
Private m_oThemeColor                           As OLE_COLOR
Private m_oGridLineColor                        As OLE_COLOR
Private m_oCellSelectedColor                    As OLE_COLOR
Private m_oCellFocusedColor                     As OLE_COLOR
Private m_oCellFocusedHighlight                 As OLE_COLOR
Private m_oHotTrackColor                        As OLE_COLOR
Private m_oFilterOffsetColor                    As OLE_COLOR
Private m_oFilterForeColor                      As OLE_COLOR
Private m_oFilterControlColor                   As OLE_COLOR
Private m_oFilterControlForeColor               As OLE_COLOR
Private m_oFilterBackColor                      As OLE_COLOR
Private c_ColumnTags                            As Collection
Private c_PtrMem                                As Collection
Private m_tDividerRect(1)                       As RECT
Private m_tREditCrd                             As RECT
Private m_oCellTipFont                          As StdFont
Private m_oFont                                 As StdFont
Private m_oHeaderFont                           As StdFont
Private m_oCellFont()                           As StdFont
Private m_pISelectorBar                         As StdPicture
Private m_IChecked                              As StdPicture
Private m_IChkDisabled                          As StdPicture
Private m_IUnChecked                            As StdPicture
Private m_cSelectorBar                          As clsStoreDc
Private m_cChkCheckDc                           As clsStoreDc
Private m_cChkUnCheckDc                         As clsStoreDc
Private m_cChkDisableDc                         As clsStoreDc
Private m_cGridBuffer                           As clsStoreDc
Private m_cTransitionMask                       As clsStoreDc
Private m_cControlDc()                          As clsStoreDc
Private m_cSizerDc                              As clsStoreDc
Private m_cSkinHeader                           As clsSkinHeader
Private m_cRender                               As clsRender
Private m_cDragImage                            As clsImageDrag
Private m_cSkinScrollBars                       As clsSkinScrollbars
Private m_cCellTips                             As clsToolTip
Private m_cHeaderIcons                          As clsImageList
Private m_cCellIcons                            As clsImageList
Private m_cTreeIcons                            As clsImageList
Private m_cSubCell()                            As clsSubCell
Private m_cCellHeader()                         As clsCellHeader
Private m_cGridItem()                           As clsGridItem
Private WithEvents m_cFilterMenu                As clsFilterMenu
Attribute m_cFilterMenu.VB_VarHelpID = -1
Private WithEvents m_cEditor                    As clsAdvancedEdit
Attribute m_cEditor.VB_VarHelpID = -1
Private WithEvents m_cEditBox                   As clsODControl
Attribute m_cEditBox.VB_VarHelpID = -1
Private WithEvents m_cTreeView                  As clsTreeView
Attribute m_cTreeView.VB_VarHelpID = -1
Private m_cHGridSubclass                        As GXMSubclass


'/~ 1.0.1 ToDo ~

'/~ Phase I
'/~ owner drawn                 - done
'/~ seperate components         - done
'/~ grid lines                  - done
'/~ cell fonts                  - done
'/~ color cells                 - done
'/~ focus effects               - done
'/~ checkboxes                  - done
'/~ header height               - done
'/~ row heights                 - done
'/~ column insert arrows        - done
'/~ hot tracking                - done
'/~ precision draw              - done
'/~ column vertical text        - done
'/~ column font                 - done
'/~ column lock                 - done
'/~ column/row focus            - done
'/~ od column sort icon         - done
'/~ column drag divider         - done
'/~ column font effects         - done
'/~ column icons                - done
'/~ xp headers                  - done
'/~ row decoration              - done
'/~ icon position               - done
'/~ ghosting                    - done
'/~ row drag and drop           - done
'/~ extend sort options         - done
'/~ sub row sorting             - done
'/~ cell key nav                - done
'/~ xp colors                   - done
'/~ header/row tooltips         - done
'/~ header height user size     - done
'/~ double buffering            - done
'/~ combo/chk/cmd/opt/list      - done
'/~ column filter               - done
'/~ custom cursors              - done
'/~ horz cell spanning          - done
'/~ vert row spanning           - done
'/~ scroll button glyph         - done
'/~ advanced edit dialog        - done
'/~ system iml connector        - done
'/~ new themes                  - done
'/~ error handling              - done
'/~ find string                 - done
'/~ subcell edit tools          - done
'/~ subcell controls            - done
'/~ cell headers                - done
'/~ unicode support             - done
'/~ virtual mode                - done
'/~ events                      - done
'/~ documentation               - done
'/~ odcell interface            - done
'/~ example projects            - done
'/~ switchable edit box style   - done

'/~ Phase II
'/~ integrated treeview         - done
'/~ tree/grid sizer bar         - done
'/~ custom draw treeview        - done
'/~ internal 32b imagelist      - done
'/~ right to left font          - done
'/~ left align scrollbars       - done
'/~ movable treeview            - done
'/~ node-to-cell ole interop    -


'/~ Bug Track ver. 1.0.1 (Pre-Release)~
'-> Column fit text             - fixed
'-> skin header reset           - fixed
'-> header pos change           - fixed
'-> sorting broken              - fixed
'-> focus text alignment        - fixed
'-> header drag arrow coord     - fixed
'-> header sizing mask          - fixed
'-> last column image           - fixed
'-> replace lvm in skinheader   - fixed
'-> header mouseover a/ resize  - fixed
'-> column locked cursor        - fixed
'-> broken by row spanning *->
'-> row count                   - fixed
'-> cell focus                  - fixed
'-> full row select             - fixed
'-> sorting                     - fixed
'-> cell navigation             - fixed
'-> cell span focus click       - fixed
'-> tooltips                    - fixed
'-> drag and drop               - fixed
'-> cell tracking               - fixed
'<-*
'-> filter bg on focused row    - fixed
'-> alpha image distortion      - fixed
'-> edit box show               - fixed
'-> clean up label tips         - fixed
'-> tooltip sorted icon         - fixed
'-> reduce var count            - fixed
'-> icon sorting broken         - fixed
'-> first column lock move      - fixed
'-> row tip off client show     - fixed
'/~ scroll btn press flicker    - fixed
'-> text-only focus s-row       - fixed
'-> checkbox hit test misfire   - fixed
'-> cell span column reorder    - fixed
'-> column flicker w/ filterbox - fixed
'-> header/list nc bg repaint   - fixed
'-> header drag image frozen    - fixed
'-> filter menu z-order         - fixed


'/~ Bug Track ver. 1.1.0, Feb 20/07 ~
'-> column resize flicker       - fixed
'-> project compile hang        - fixed
'-> grid pre-cell focus         - fixed
'-> focus text w/ cell header   - fixed
'-> checkbox hittest error      - fixed
'-> header cursor change        - fixed
'-> sort griditem distortion    - fixed
'-> add/remove column fail      - fixed

'/~ Bug Track ver. 1.2.0, Feb 23/07 ~
'-> property settings           - fixed
'-> font change                 - fixed
'-> header hide broken          - fixed
'-> header extended length      - fixed

'/~ Bug Track ver. 1.3.0, Feb 26/07 ~
'-> initial cell count size     - fixed
'-> listbox focus rect          - fixed
'-> post filter row count       - fixed
'-> add/remove/clear index      - fixed
'-> editbox cell focus          - fixed
'-> header hit-testing          - fixed
'-> cursor changes              - fixed
'-> header size while dragging  - fixed
'-> row select focus color      - fixed
'-> enum/constants cleanup      - fixed

'/~ Bug Track ver. 1.4.0, Mar 04/07 ~
'-> horz scrlbar btn dissapear  - fixed
'-> bg change refresh           - fixed
'-> filter unload refresh       - fixed
'-> od example font color       - fixed
'-> settings and display change - fixed
'-> transmask refresh when szg  - fixed

'/~ Bug Track ver. 1.5.0, Mar 18/07 ~
'-> frame scroll button         - fixed
'-> visible toggle scrollbars   - fixed
'-> timer message release       - fixed
'-> clip text within cell       - fixed

'/~ Bug Track ver. 1.6.0, Mar 25/07 ~
'-> edit close m/wheel scroll   - fixed
'-> checkbox not sorting        - fixed
'-> columns not fully locked    - fixed
'-> rem column w/ cell span     - fixed



'**********************************************************************
'*                            EVENTS
'**********************************************************************

Private Sub m_cFilterMenu_DestroyMe()

    Set m_cFilterMenu = Nothing
    GridRefresh False
    FilterLoaded = False

End Sub

Private Sub m_cFilterMenu_FilterIndex(ByVal lIndex As Long)

    m_bSorted = False
    If (lIndex = 0) Then
        If Not (m_lRowCount = 0) Then
            m_bFiltered = False
            FilterReset
        End If
    Else
        If Not (m_lRowCount = 0) Then
            If m_bFilterHeader Then
                FilterApplyHeader lIndex
            Else
                FilterApply lIndex
            End If
            m_bFiltered = True
        End If
    End If

End Sub

Private Sub m_cEditBox_LostFocus()
    GridFocus = True
End Sub

Private Sub m_cEditor_ReturnData(ByVal sText As String, _
                                 ByVal oFont As stdole.StdFont, _
                                 ByVal lIcon As Long, _
                                 ByVal lForeColor As Long, _
                                 ByVal lBackColor As Long)

    If (m_lEditItem > -1) Then
        If (m_lEditSubItem > -1) Then
            If m_bVirtualMode Then
                RaiseEvent eVHAdvancedEditChange(m_lEditItem, m_lEditSubItem, lIcon, lBackColor, lForeColor, sText)
            Else
                With m_cGridItem(m_lEditItem)
                    .Text(m_lEditSubItem) = sText
                    If Not (oFont Is Nothing) Then
                        .FontHnd(m_lEditSubItem) = CellAddFont(oFont)
                    End If
                    If Not (lIcon = -1) Then
                        .Icon(m_lEditSubItem) = lIcon
                    End If
                    If Not (lForeColor = -1) Then
                        .ForeColor(m_lEditSubItem) = lForeColor
                    End If
                    If Not (lBackColor = -1) Then
                        .BackColor(m_lEditSubItem) = lBackColor
                    End If
                End With
            End If
            GridRefresh False
        End If
    End If

End Sub

Private Sub m_cEditor_DestroyMe()

    Set m_cEditor = Nothing
    m_bEditorLoaded = False
    m_lEditItem = -1
    m_lEditSubItem = -1

End Sub

Private Sub m_cTreeView_AfterExpand(ByVal hNode As Long, ByVal ExpandedOnce As Boolean)
    RaiseEvent eTVAfterExpand(hNode, ExpandedOnce)
End Sub

Private Sub m_cTreeView_AfterLabelEdit(ByVal hNode As Long, Cancel As Long, NewString As String)
    RaiseEvent eTVAfterLabelEdit(hNode, Cancel, NewString)
End Sub

Private Sub m_cTreeView_BeforeExpand(ByVal hNode As Long, ByVal ExpandedOnce As Boolean)
    RaiseEvent eTVBeforeExpand(hNode, ExpandedOnce)
End Sub

Private Sub m_cTreeView_BeforeLabelEdit(ByVal hNode As Long, Cancel As Long)
    RaiseEvent eTVBeforeLabelEdit(hNode, Cancel)
End Sub

Private Sub m_cTreeView_Click()
    RaiseEvent eTVClick
End Sub

Private Sub m_cTreeView_Collapse(ByVal hNode As Long)
    RaiseEvent eTVCollapse(hNode)
End Sub

Private Sub m_cTreeView_KeyDown(KeyCode As Long, Shift As Long)
    RaiseEvent eTVKeyDown(KeyCode, Shift)
End Sub

Private Sub m_cTreeView_KeyPress(KeyAscii As Long)
    RaiseEvent eTVKeyPress(KeyAscii)
End Sub

Private Sub m_cTreeView_KeyUp(KeyCode As Long, Shift As Long)
    RaiseEvent eTVKeyUp(KeyCode, Shift)
End Sub

Private Sub m_cTreeView_MouseDown(Button As Long, Shift As Long, x As Long, y As Long)
    RaiseEvent eTVMouseDown(Button, Shift, x, y)
End Sub

Private Sub m_cTreeView_MouseEnter()
    RaiseEvent eTVMouseEnter
End Sub

Private Sub m_cTreeView_MouseLeave()
    RaiseEvent eTVMouseLeave
End Sub

Private Sub m_cTreeView_MouseMove(Button As Long, Shift As Long, x As Long, y As Long)
    RaiseEvent eTVMouseMove(Button, Shift, x, y)
End Sub

Private Sub m_cTreeView_MouseUp(Button As Long, Shift As Long, x As Long, y As Long)
    RaiseEvent eTVMouseUp(Button, Shift, x, y)
End Sub

Private Sub m_cTreeView_NodeCheck(ByVal hNode As Long)
    RaiseEvent eTVNodeCheck(hNode)
End Sub

Private Sub m_cTreeView_NodeClick(ByVal hNode As Long)
    RaiseEvent eTVNodeClick(hNode)
End Sub

Private Sub m_cTreeView_NodeDblClick(ByVal hNode As Long)
    RaiseEvent eTVNodeDblClick(hNode)
End Sub

Private Sub m_cTreeView_SelectionChanged()
    RaiseEvent eTVSelectionChanged
End Sub


'**********************************************************************
'*                            CONSTRUCTORS
'**********************************************************************

Private Sub UserControl_Initialize()

'/* init control

'Dim i As Long

' i = VarPtr(m_lEditItem) Mod 8
' If i Then
'     Debug.Print "Misaligned by " & i
' End If

    m_lhMod = LoadLibrary("shell32.dll")
    InitCommonControls
    InitComctl32
    VersionCheck
    Set m_cHGridSubclass = New GXMSubclass
    m_oBackColor = &HFFFFFF
    m_oThemeColor = -1
    m_oGridLineColor = GetSysColor(vbButtonShadow And &H1F&)
    m_oCellSelectedColor = GetSysColor(vbButtonFace And &H1F&)
    m_oCellFocusedColor = GetSysColor(vbHighlight And &H1F&)
    m_oCellFocusedHighlight = GetSysColor(vbHighlightText And &H1F&)
    m_oColumnFocusColor = &H303030
    m_lDisabledBackColor = &HCCCCCC
    m_lDisabledForeColor = &H999999
    m_bEnabled = True
    m_eCheckBoxSkinStyle = evsXpSilver
    BorderStyle = ebsThin
    m_oHotTrackColor = GetSysColor(vbHighlight And &H1F&)
    m_eHotTrackDepth = ehtNarrow
    m_lRowHeight = 20
    m_lRowMinHgt = 20
    m_lHeaderHeight = 24
    m_lColumnSelected = -1
    m_oCellTipColor = GetSysColor(&H80000018 And &H1F)
    m_oCellTipOffsetColor = &HDEDEDE
    m_lCellTipDelayTime = 2
    m_oCellTipForeColor = &H313131
    m_bCellTipMultiline = True
    m_lCellTipVisibleTime = 3
    m_lCellTipTransparency = -1
    m_lAdvancedEditThemeColor = -1
    m_lAdvancedEditOffsetColor = -1
    m_lCellHiliteColor = &H666666
    ReDim m_sFilterItem(0)
    Set c_ColumnTags = New Collection
    Set m_oFont = New StdFont
    Set m_cRender = New clsRender
    Set m_cDragImage = New clsImageDrag
    Set m_cSkinHeader = New clsSkinHeader
    Set m_cHeaderIcons = New clsImageList
    Set m_cCellIcons = New clsImageList
    Set m_cTreeIcons = New clsImageList
'    Stop
    ImlUseAlphaIcons = True
    
End Sub

Public Function CreateGrid() As Boolean
'*/ constructor: initialize the grid

Dim lLVStyle As Long
Dim lExStyle As Long
Dim tRect    As RECT

    '/* destroy existing
    DestroyList
    m_lParentHwnd = UserControl.hwnd
    GetClientRect m_lParentHwnd, tRect
    '/* initial style flags including LVS_OWNERDATA
    '/* this tells the list that all data will be
    '/* managed externally
    lLVStyle = WS_CHILD Or WS_BORDER Or WS_VISIBLE Or WS_TABSTOP Or LVS_SORTASCENDING Or _
        LVS_OWNERDATA Or LVS_SHAREIMAGELISTS Or LVS_SHOWSELALWAYS Or LVS_SINGLESEL Or _
        LVS_REPORT Or LVS_OWNERDRAWFIXED Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN
    If m_bFontRightLeading Then
        lExStyle = WS_EX_RTLREADING
    End If
    '/* create listview
    If m_bIsNt Then
        With tRect
            m_lHGHwnd = CreateWindowExW(lExStyle, StrPtr(WC_LISTVIEW), StrPtr(""), lLVStyle, 0&, 0&, (.Right - .left), (.Bottom - .top), m_lParentHwnd, 0&, App.hInstance, ByVal 0&)
        End With
    Else
        With tRect
            m_lHGHwnd = CreateWindowExA(0&, WC_LISTVIEW, "", lLVStyle, 0&, 0&, .Right - .left, .Bottom - .top, m_lParentHwnd, 0&, App.hInstance, ByVal 0&)
        End With
    End If
    m_lHdrHwnd = HeaderHwnd
    SetUnicode True
    If m_bIsXp Then
        SetWindowTheme m_lHGHwnd, StrPtr(" "), StrPtr(" ")
        SetWindowTheme m_lHdrHwnd, StrPtr(" "), StrPtr(" ")
    End If
    '/* default border style
    SetBorderStyle m_lHGHwnd, ebsThin
    '/* subclass the list and parent WM_NOTIFY messages
    '/* control callback data is reflected from parent control
    If Not (m_lHGHwnd = 0) Then
        GridAttatch
        m_cSkinHeader.ParentHwnd = m_lHGHwnd
    End If

End Function

Public Property Get GetImlObj(Optional ByVal sPropName As String) As Long
Attribute GetImlObj.VB_MemberFlags = "40"
    If (sPropName = "HeaderIcons") Then
        GetImlObj = ObjPtr(m_cHeaderIcons)
    ElseIf (sPropName = "CellIcons") Then
        GetImlObj = ObjPtr(m_cCellIcons)
    ElseIf (sPropName = "TreeIcons") Then
        GetImlObj = ObjPtr(m_cTreeIcons)
    End If
End Property

Public Property Get HeaderIcons() As Long
Attribute HeaderIcons.VB_ProcData.VB_Invoke_Property = "ppgImages"
    If Not (m_cHeaderIcons Is Nothing) Then
        HeaderIcons = m_cHeaderIcons.ImageCount
    End If
End Property

Public Property Let HeaderIcons(ByVal PropVal As Long)
    If Not (m_cHeaderIcons Is Nothing) Then
        m_cHeaderIcons.ImageCount = PropVal
        PropertyChanged "HeaderIconCount"
    End If
End Property

Public Property Get CellIcons() As Long
Attribute CellIcons.VB_ProcData.VB_Invoke_Property = "ppgImages"
    If Not (m_cCellIcons Is Nothing) Then
        CellIcons = m_cCellIcons.ImageCount
    End If
End Property

Public Property Let CellIcons(ByVal PropVal As Long)
    If Not (m_cCellIcons Is Nothing) Then
        m_cCellIcons.ImageCount = PropVal
        PropertyChanged "CellIconCount"
    End If
End Property

Public Property Get TreeIcons() As Long
Attribute TreeIcons.VB_ProcData.VB_Invoke_Property = "ppgImages"
    If Not (m_cTreeIcons Is Nothing) Then
        TreeIcons = m_cTreeIcons.ImageCount
    End If
End Property

Public Property Let TreeIcons(ByVal PropVal As Long)
    If Not (m_cTreeIcons Is Nothing) Then
        m_cTreeIcons.ImageCount = PropVal
        PropertyChanged "TreeIconCount"
    End If
End Property

Public Function CellIconAdd(ByVal lhImage As Long, _
                            Optional ByVal sKey As String, _
                            Optional ByVal lMask As Long = -1) As Boolean
    If Not (m_cCellIcons Is Nothing) Then
        CellIconAdd = m_cCellIcons.AddFromHandle(lhImage, IMAGE_ICON, sKey, lMask)
    End If
End Function

Public Sub CellIconRemove(ByVal vIcon As Variant)
    If Not (m_cCellIcons Is Nothing) Then
        m_cCellIcons.RemoveImage vIcon
    End If
End Sub

Public Sub CellIconsClear()
    If Not (m_cCellIcons Is Nothing) Then
        m_cCellIcons.Clear
    End If
End Sub

Public Function CellIconIndexFromKey(ByVal sKey As String) As Long
    If Not (m_cCellIcons Is Nothing) Then
        CellIconIndexFromKey = m_cCellIcons.IndexFromKey(sKey)
    End If
End Function

Public Function CellIconKeyFromIndex(ByVal lIndex As Long) As String
    If Not (m_cCellIcons Is Nothing) Then
        CellIconKeyFromIndex = m_cCellIcons.KeyFromIndex(lIndex)
    End If
End Function

Public Function CellIconKeyExists(ByVal sKey As String) As Boolean
    If Not (m_cCellIcons Is Nothing) Then
        CellIconKeyExists = m_cCellIcons.KeyExists(sKey)
    End If
End Function

Public Function CellIconAddIconFromFileName(ByVal sFile As String, _
                                            Optional ByVal sKey As String) As Boolean

Dim lSize   As Long
Dim lhIcon  As Long
Dim lSzFlag As Long

    If Not (m_cCellIcons Is Nothing) Then
        If Not (Len(sFile) = 0) Then
            lSize = m_cCellIcons.IconSizeX
            Select Case lSize
            Case 16
                lSzFlag = SHGFI_SHELLICONSIZE Or SHGFI_SMALLICON
            Case 24
                lSzFlag = SHGFI_ICON Or SHGFI_SMALLICON
            Case 32
                lSzFlag = SHGFI_ICON Or SHGFI_LARGEICON
            Case 48
                lSzFlag = SHGFI_SHELLICONSIZE Or SHGFI_ICON
            Case Else
                lSzFlag = SHGFI_ICON Or SHGFI_LARGEICON
            End Select
            lhIcon = m_cCellIcons.SystemIconHandle(sFile, lSzFlag)
            If Not (lhIcon = 0) Then
                m_cCellIcons.AddFromHandle lhIcon, IMAGE_ICON, sKey
            End If
        End If
    End If
    
End Function

Private Function GetThemeName() As EXTXpTheme

Dim bClrNm()    As Byte
Dim bTmeNm()    As Byte
Dim lhTheme     As Long
Dim lPtrTme     As Long
Dim lPtrClr     As Long
Dim sClass      As String
Dim sTheme      As String

On Error GoTo Handler

    sClass = "listview"
    lhTheme = OpenThemeData(m_lParentHwnd, StrPtr(sClass))
    If lhTheme = 0 Then GoTo Handler
    
    ReDim bTmeNm(0 To 260 * 2) As Byte
    lPtrTme = VarPtr(bTmeNm(0))
    ReDim bClrNm(0 To 260 * 2) As Byte
    lPtrClr = VarPtr(bClrNm(0))
    GetCurrentThemeName lPtrTme, 260&, lPtrClr, 260&, 0&, 0&
    sTheme = bClrNm
    If Not InStr(sTheme, vbNullChar) = 0 Then
        sTheme = left$(sTheme, InStr(sTheme, vbNullChar) - 1)
    End If
    
    Select Case sTheme
    Case "HomeStead"
        GetThemeName = HomeStead
    Case "NormalColor"
        GetThemeName = NormalColor
    Case "Metallic"
        GetThemeName = Metallic
    Case Else
        GetThemeName = None
    End Select
    
On Error GoTo 0

Handler:
    If Not (lhTheme = 0) Then
        CloseThemeData lhTheme
    End If
    sTheme = "None"

End Function

Public Sub GridInit(ByVal lRowCount As Long, _
                    ByVal lCellCount As Long)

'/* initialize grid arrays

Dim lCt As Long

    If Not (lRowCount = 0) Then
        If Not (lCellCount = 0) Then
            lRowCount = lRowCount - 1
            lCellCount = lCellCount - 1
            If Not m_bFastLoad Then
                m_bDraw = True
            End If
            If Not m_bVirtualMode Then
                ReDim m_cGridItem(0 To lRowCount)
                For lCt = 0 To lRowCount
                    Set m_cGridItem(lCt) = New clsGridItem
                    m_cGridItem(lCt).Init lCellCount
                Next lCt
            End If
            m_lRowCount = lRowCount + 1
            SetRowCount (m_lRowCount)
        End If
    End If

End Sub

Public Property Get hwnd() As Long
Attribute hwnd.VB_MemberFlags = "40"
    hwnd = m_lHGHwnd
End Property

Private Function InitComctl32() As Boolean
'/* init comctl32 listview class

Dim tIcc As tagINITCOMMONCONTROLSEX

    With tIcc
        .dwSize = Len(tIcc)
        .dwICC = ICC_LISTVIEW_CLASSES
    End With
    InitComctl32 = InitCommonControlsEx(tIcc)

End Function

Public Function IsUnicode(ByVal sText As String) As Boolean
'/* test for unicode
'/* good link: http://www.unicodeactivex.com/UnicodeTutorialVb.htm

Dim lLen   As Long
Dim bLen   As Long
Dim bmap() As Byte

    If LenB(sText) Then
        bmap = sText
        bLen = UBound(bmap)
        For lLen = 1 To bLen Step 2
            If (bmap(lLen) > 0) Then
                IsUnicode = True
                Exit For
            End If
        Next lLen
    End If

End Function

Private Function LeftKeyState() As Boolean
'/* left button pressed state

    If ((GetKeyState(VK_LBUTTON) And &H80) > 1) Then
        LeftKeyState = True
    End If

End Function

Public Function LoadArray() As Boolean
'*/ load data structure

    If Not m_bVirtualMode Then
        Set c_PtrMem = New Collection
        '/* initialize local struct
        ReDim m_cGridItem(0)
        '/* copy the structure from the pointer
        CopyMemory ByVal VarPtrArray(m_cGridItem), m_lStrctPtr, 4&
        c_PtrMem.Add m_lStrctPtr, "m_cGridItem"
        m_lRowCount = UBound(m_cGridItem) + 1
        LoadArray = True
    End If

End Function

Private Sub Initialize()
'/* pre initialize grid

Dim bRun As Boolean

    bRun = UserControl.Ambient.UserMode
    If bRun Then
        CreateGrid
    End If

End Sub

Private Function SetUnicode(ByVal bEnable As Boolean) As Boolean
'/* enable/disable unicode processing

Dim lRet As Long

    If Not (m_lHGHwnd = 0) Then
        If m_bIsNt Then
            If bEnable Then
                If Not UnicodeState Then
                    lRet = SendMessageLongW(m_lHGHwnd, CCM_SETUNICODEFORMAT, 1&, 0&)
                End If
            Else
                If UnicodeState Then
                    lRet = SendMessageLongW(m_lHGHwnd, CCM_SETUNICODEFORMAT, 0&, 0&)
                End If
            End If
        End If
    Else
        lRet = -1
    End If
    SetUnicode = (lRet = 0)

End Function

Private Function UnicodeState() As Boolean
'/* get control unicode readiness

    If Not (m_lHGHwnd = 0) Then
        UnicodeState = SendMessageLongW(m_lHGHwnd, CCM_GETUNICODEFORMAT, 0&, 0&) <> 0
    End If

End Function

Private Function PointerToString(ByVal lpString As Long) As String
'/* get string from pointer

Dim lLen As Long

    If Not (lpString = 0) Then
        lLen = lstrlenW(ByVal lpString)
        If Not (lLen = 0) Then
            '/* allocate string with nLen chars
            PointerToString = String$(lLen, Chr$(0))
            lstrcpyW StrPtr(PointerToString), lpString
        End If
    End If

End Function

Private Function VersionCheck() As Boolean
'/* operating system check

Dim tVer As OSVERSIONINFO

    With tVer
        .dwVersionInfoSize = LenB(tVer)
        GetVersionEx tVer
        m_bIsNt = ((.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
        If m_bIsNt Then
            If (.dwMajorVersion >= 5) Then
                m_bIsXp = True
            End If
        End If
    End With
    If Not m_bIsNt Then
        m_bUseUnicode = False
    End If
    VersionCheck = m_bIsNt

End Function


'**********************************************************************
'*                            TREEVIEW
'**********************************************************************

Private Sub DrawSizer(ByVal lHdc As Long)
'/* draw treeview sizer bar

Dim lCt         As Long
Dim lFillStyle  As Long
Dim lHeight     As Long
Dim tRect       As RECT

    Select Case m_eTvControlAlignment
    Case etvLeftAlign
        lFillStyle = 0
        GetWindowRect m_lTVHwnd, tRect
        With tRect
            OffsetRect tRect, -.left, -.top
            .left = .Right
            .Right = .left + 10
            lHeight = (.Bottom + 30)
        End With
        With m_cSizerDc
            .Width = 10
            .Height = tRect.Bottom
        End With
    Case etvRightAlign
        lFillStyle = 0
        GetWindowRect m_lHGHwnd, tRect
        With tRect
            OffsetRect tRect, -.left, -.top
            .left = .Right
            .Right = .left + 10
            lHeight = (.Bottom + 30)
        End With
        With m_cSizerDc
            .Width = 10
            .Height = tRect.Bottom
        End With
    Case etvTopAlign
        lFillStyle = 1
        GetWindowRect m_lTVHwnd, tRect
        With tRect
            OffsetRect tRect, -.left, -.top
            .top = .Bottom
            .Bottom = .top + 10
            lHeight = (.Right + 40)
        End With
        With m_cSizerDc
            .Width = tRect.Right
            .Height = 10
        End With
    Case etvBottomAlign
        lFillStyle = 1
        GetWindowRect m_lHGHwnd, tRect
        With tRect
            OffsetRect tRect, -.left, -.top
            .top = .Bottom
            .Bottom = .top + 10
            lHeight = (.Right + 40)
        End With
        With m_cSizerDc
            .Width = tRect.Right
            .Height = 10
        End With
    End Select
    
    With m_cSizerDc
        m_cRender.DrawRectangle .hdc, 0&, 0&, .Width, .Height, &H666666
        m_cRender.Gradient .hdc, 1, .Width - 2, 1, .Height - 2, m_oBackColor, m_cRender.BlendColor(m_oBackColor, &HB3B3B3), lFillStyle
        If m_bTreeViewSizeable Then
            If (m_eTvControlAlignment > 1) Then
                For lCt = lCt To 30 Step 5
                    SetPixel .hdc, (lHeight \ 2) - (2 + lCt), 5, vbBlack
                    SetPixel .hdc, (lHeight \ 2) - (1 + lCt), 5, vbBlack
                    SetPixel .hdc, (lHeight \ 2) - (1 + lCt), 6, vbWhite
                Next lCt
            Else
                For lCt = lCt To 30 Step 5
                    SetPixel .hdc, 5, (lHeight \ 2) - (2 + lCt), vbBlack
                    SetPixel .hdc, 5, (lHeight \ 2) - (1 + lCt), vbBlack
                    SetPixel .hdc, 6, (lHeight \ 2) - (1 + lCt), vbWhite
                Next lCt
            End If
        End If
    End With
    With tRect
        m_cRender.Blit lHdc, .left, .top, .Right, .Bottom, m_cSizerDc.hdc, 0, 0, SRCCOPY
    End With

End Sub

Public Sub TreeViewInit(ByVal lWidth As Long, _
                        ByVal lHeight As Long, _
                        Optional ByVal lBackColor As Long = -1, _
                        Optional ByVal lForeColor As Long = -1, _
                        Optional ByVal lLineColor As Long = &H343434, _
                        Optional ByVal lItemHeight As Long = 18, _
                        Optional ByVal lItemIndent As Long = 4, _
                        Optional ByVal bCheckBoxes As Boolean, _
                        Optional ByVal bSizeable As Boolean, _
                        Optional ByVal bNodeLines As Boolean, _
                        Optional ByVal bRootLines As Boolean, _
                        Optional ByVal bButtons As Boolean, _
                        Optional ByVal bFullRowSelect As Boolean, _
                        Optional ByVal bLabelEdit As Boolean, _
                        Optional ByVal bTrackSelect As Boolean, _
                        Optional ByVal bScrollBarLeft As Boolean)
'/* initialize treeview

Dim tRect As RECT

    Set m_cTreeView = New clsTreeView
    GetClientRect m_lParentHwnd, tRect
    With m_cTreeView
        Select Case m_eTvControlAlignment
        Case etvLeftAlign, etvRightAlign
            .Initialize UserControl.hwnd, tRect.Bottom, lWidth
        Case etvTopAlign, etvBottomAlign
            .Initialize UserControl.hwnd, lHeight, tRect.Right
        End Select
        .FontRightLeading = m_bFontRightLeading
        .GridHwnd = m_lHGHwnd
        If (lBackColor = -1) Then
            If m_bXPColors Then
                .BackColor = m_cRender.XPShift(m_oBackColor)
            Else
                .BackColor = m_oBackColor
            End If
        Else
            If m_bXPColors Then
                .BackColor = m_cRender.XPShift(lBackColor)
            Else
                .BackColor = lBackColor
            End If
        End If
        .BorderStyle = tbsNone
        If bCheckBoxes Then
            .CheckBoxes = True
            .SkinCheckBox m_eCheckBoxSkinStyle
        End If
        Set .Font = m_oFont
        If (lForeColor = -1) Then
            .ForeColor = m_oForeColor
        Else
            .ForeColor = lForeColor
        End If
        .DisabledBackColor = m_lDisabledBackColor
        .DisabledForeColor = m_lDisabledForeColor
        .FocusBackColor = m_oCellFocusedColor
        .FocusForeColor = m_oCellFocusedHighlight
        .SelectedBackColor = m_oCellSelectedColor
        .FullRowSelect = bFullRowSelect
        .HasButtons = bButtons
        .HasLines = bNodeLines
        .HasRootLines = bRootLines
        .ItemHeight = lItemHeight
        .ItemIndent = lItemIndent
        .LabelEdit = bLabelEdit
        .LineColor = lLineColor
        .TrackSelect = bTrackSelect
        .UseUnicode = m_bUseUnicode
        If (m_eTvControlAlignment = etvRightAlign) Then
            .ScrollBarLeftAlign = True
            ScrollBarAlignment = escLeftAlign
        Else
            .ScrollBarLeftAlign = bScrollBarLeft
        End If
        .SkinScrollBars m_eScrollBarSkinStyle
        m_lTVHwnd = .hwnd
    End With
    m_bHasTreeView = True
    m_bTreeViewSizeable = bSizeable
    TreeViewAddSizer
    Resize

End Sub

Public Function TreeViewInitIml()
'/* init treeview iml

    If Not (m_cTreeView Is Nothing) Then
        With m_cTreeView
           .hImlNode = m_cTreeIcons.hIml
            If .CheckBoxes Then
                .SkinCheckBox m_eCheckBoxSkinStyle
            End If
        End With
    End If

End Function

Public Function TreeViewAddNode(Optional ByVal lRelative As Long, _
                                Optional ByVal eRelation As ETVNodeRelation, _
                                Optional ByVal sKey As String, _
                                Optional ByVal sText As String, _
                                Optional ByVal lImage As Long = -1, _
                                Optional ByVal lSelectedImage As Long = -1, _
                                Optional ByVal bPlusButton As Boolean = False, _
                                Optional ByVal sTag As String = vbNullString) As Long
'/* add nodes to treeview

    If Not (m_cTreeView Is Nothing) Then
        With m_cTreeView
            TreeViewAddNode = .AddNode(lRelative, eRelation, sKey, sText, lImage, lSelectedImage, bPlusButton, sTag)
        End With
    End If
    
End Function

Public Sub TreeViewClear()
'/* clear treeview

    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.Clear
    End If

End Sub

Public Sub TreeViewCheckChildren(ByVal lNode As Long, _
                                 ByVal bChecked As Boolean)
'/* check node children

    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.CheckChildren lNode, bChecked
    End If

End Sub

Public Sub TreeViewCollapse(ByVal lNode As Long, _
                            ByVal bChildren As Boolean)
'/* collapse a node

    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.Collapse lNode, bChildren
    End If

End Sub

Public Sub TreeViewDraw(ByVal bDraw As Boolean)
'/* treeview draw switch

    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.SetRedrawMode bDraw
    End If

End Sub

Public Sub TreeViewEnsureVisible(ByVal lNode As Long)
'/* node ensure visible
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.EnsureVisible lNode
    End If

End Sub

Public Sub TreeViewExpand(ByVal lNode As Long, _
                          ByVal bChildren As Boolean)
'/* expand a node

    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.Expand lNode, bChildren
    End If
    
End Sub

Public Function TreeViewGetKeyNode(ByVal sKey As String) As Long
'/* get node key index

    If Not (m_cTreeView Is Nothing) Then
        TreeViewGetKeyNode = m_cTreeView.GetKeyNode(sKey)
    End If
    
End Function

Public Function TreeViewGetNodeKey(ByVal lNode As Long) As String
'/* fetch node key

    If Not (m_cTreeView Is Nothing) Then
        TreeViewGetNodeKey = m_cTreeView.GetNodeKey(lNode)
    End If
    
End Function

Public Sub TreeViewSortChildren(ByVal lNode As Long, _
                                ByVal bAllLevels As Boolean)
Attribute TreeViewSortChildren.VB_MemberFlags = "40"
'/* sort treeview

    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.SortChildren lNode, bAllLevels
    End If
    
End Sub

Private Sub TreeViewAddSizer()
'/* add treeview sizer messages

    If m_bHasTreeView Then
        If Not (m_cTreeView Is Nothing) Then
            Set m_cSizerDc = New clsStoreDc
            If Not (m_lParentHwnd = 0) Then
                With m_cHGridSubclass
                    .AddMessage m_lParentHwnd, WM_PAINT, MSG_BEFORE
                    .AddMessage m_lParentHwnd, WM_ERASEBKGND, MSG_BEFORE
                    .AddMessage m_lParentHwnd, WM_SETCURSOR, MSG_BEFORE
                    If Not (m_lTVHwnd = 0) Then
                        .Subclass m_lTVHwnd, Me
                        .AddMessage m_lTVHwnd, WM_SETCURSOR, MSG_BEFORE
                    End If
                End With
            End If
        End If
    End If

End Sub

Private Sub TreeRemoveSizer()
'/* remove sizer subclass messages

    If m_bHasTreeView Then
        If Not (m_cTreeView Is Nothing) Then
            If Not (m_lParentHwnd = 0) Then
                With m_cHGridSubclass
                    .DeleteMessage m_lParentHwnd, WM_PAINT, MSG_BEFORE
                    .DeleteMessage m_lParentHwnd, WM_ERASEBKGND, MSG_BEFORE
                    .DeleteMessage m_lParentHwnd, WM_SETCURSOR, MSG_BEFORE
                    If Not (m_lTVHwnd = 0) Then
                        .DeleteMessage m_lTVHwnd, WM_SETCURSOR, MSG_BEFORE
                        .UnSubclass m_lTVHwnd
                    End If
                End With
            End If
        End If
    End If

End Sub

Public Property Get TreeViewAlignment() As ETVTreeViewAlignment
Attribute TreeViewAlignment.VB_MemberFlags = "40"
'/* [get] treeview control alignment
    TreeViewAlignment = m_eTvControlAlignment
End Property

Public Property Let TreeViewAlignment(ByVal PropVal As ETVTreeViewAlignment)
'/* [let] treeview control alignment
    m_eTvControlAlignment = PropVal
    If Not (m_cTreeView Is Nothing) Then
        Select Case m_eTvControlAlignment
        Case etvBottomAlign
            m_cTreeView.ScrollBarLeftAlign = True
            ScrollBarAlignment = escLeftAlign
        Case etvLeftAlign
            m_cTreeView.ScrollBarLeftAlign = True
            ScrollBarAlignment = escRightAlign
        Case etvRightAlign
            m_cTreeView.ScrollBarLeftAlign = False
            ScrollBarAlignment = escLeftAlign
        Case etvTopAlign
            m_cTreeView.ScrollBarLeftAlign = True
            ScrollBarAlignment = escLeftAlign
        End Select
        Set m_cSizerDc = Nothing
        Set m_cSizerDc = New clsStoreDc
        Resize
        GridRefresh True
        m_cTreeView.Refresh True
    End If
End Property

Public Property Get TreeViewBackColor() As Long
Attribute TreeViewBackColor.VB_MemberFlags = "40"
'/* [get] treeview backcolor
    If Not (m_cTreeView Is Nothing) Then
        TreeViewBackColor = m_cTreeView.BackColor
    End If
End Property

Public Property Let TreeViewBackColor(ByVal PropVal As Long)
'/* [let] treeview backcolor
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.BackColor = PropVal
    End If
End Property

Public Property Get TreeViewCheckBoxes() As Boolean
Attribute TreeViewCheckBoxes.VB_MemberFlags = "40"
'/* [get] treeview checkboxes
    If Not (m_cTreeView Is Nothing) Then
        TreeViewCheckBoxes = m_cTreeView.CheckBoxes
    End If
End Property

Public Property Let TreeViewCheckBoxes(ByVal PropVal As Boolean)
'/* [let] treeview checkboxes
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.CheckBoxes = PropVal
    End If
End Property

Public Property Get TreeViewEnabled() As Boolean
Attribute TreeViewEnabled.VB_MemberFlags = "40"
'/* [get] treeview enabled
    If Not (m_cTreeView Is Nothing) Then
        TreeViewEnabled = m_cTreeView.Enabled
    End If
End Property

Public Property Let TreeViewEnabled(ByVal PropVal As Boolean)
'/* [let] treeview enabled
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.Enabled = PropVal
    End If
End Property

Public Property Get TreeViewDisabledBackColor() As Long
Attribute TreeViewDisabledBackColor.VB_MemberFlags = "40"
'/* [get] treeview disabled backcolor
    If Not (m_cTreeView Is Nothing) Then
        TreeViewDisabledBackColor = m_cTreeView.DisabledBackColor
    End If
End Property

Public Property Let TreeViewDisabledBackColor(ByVal PropVal As Long)
'/* [let] treeview disabled backcolor
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.DisabledBackColor = PropVal
    End If
End Property

Public Property Get TreeViewDisabledForeColor() As Long
Attribute TreeViewDisabledForeColor.VB_MemberFlags = "40"
'/* [get] treeview disabled forecolor
    If Not (m_cTreeView Is Nothing) Then
        TreeViewDisabledForeColor = m_cTreeView.DisabledForeColor
    End If
End Property

Public Property Let TreeViewDisabledForeColor(ByVal PropVal As Long)
'/* [let] treeview disabled forecolor
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.DisabledForeColor = PropVal
    End If
End Property

Public Property Set TreeViewFont(ByVal oFont As StdFont)
Attribute TreeViewFont.VB_MemberFlags = "40"
'/* [set] treeview font
    If Not (m_cTreeView Is Nothing) Then
        Set m_cTreeView.Font = oFont
    End If
End Property

Public Property Get TreeViewFocusBackColor() As Long
Attribute TreeViewFocusBackColor.VB_MemberFlags = "40"
'/* [get] treeview focused backcolor
    If Not (m_cTreeView Is Nothing) Then
        TreeViewFocusBackColor = m_cTreeView.FocusBackColor
    End If
End Property

Public Property Let TreeViewFocusBackColor(ByVal PropVal As Long)
'/* [let] treeview focused backcolor
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.FocusBackColor = PropVal
    End If
End Property

Public Property Let TreeViewFocusForeColor(ByVal PropVal As Long)
'/* [let] treeview focused forecolor
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.FocusForeColor = PropVal
    End If
End Property

Public Property Get TreeViewFocusForeColor() As Long
Attribute TreeViewFocusForeColor.VB_MemberFlags = "40"
'/* [get] treeview focused forecolor
    If Not (m_cTreeView Is Nothing) Then
        TreeViewFocusForeColor = m_cTreeView.FocusForeColor
    End If
End Property

Public Property Get TreeViewForeColor() As Long
Attribute TreeViewForeColor.VB_MemberFlags = "40"
'/* [get] treeview forecolor
    If Not (m_cTreeView Is Nothing) Then
        TreeViewForeColor = m_cTreeView.ForeColor
    End If
End Property

Public Property Let TreeViewForeColor(ByVal PropVal As Long)
'/* [let] treeview forecolor
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.ForeColor = PropVal
    End If
End Property

Public Property Get TreeViewFullRowSelect() As Boolean
Attribute TreeViewFullRowSelect.VB_MemberFlags = "40"
'/* [get] treeview full row select
    If Not (m_cTreeView Is Nothing) Then
        TreeViewFullRowSelect = m_cTreeView.FullRowSelect
    End If
End Property

Public Property Let TreeViewFullRowSelect(ByVal PropVal As Boolean)
'/* [let] treeview full row select
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.FullRowSelect = PropVal
    End If
End Property

Public Property Get TreeViewHasButtons() As Boolean
Attribute TreeViewHasButtons.VB_MemberFlags = "40"
'/* [get] treeview buttons
    If Not (m_cTreeView Is Nothing) Then
        TreeViewHasButtons = m_cTreeView.HasButtons
    End If
End Property

Public Property Let TreeViewHasButtons(ByVal PropVal As Boolean)
'/* [let] treeview buttons
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.HasButtons = PropVal
    End If
End Property

Public Property Get TreeViewHasLines() As Boolean
Attribute TreeViewHasLines.VB_MemberFlags = "40"
'/* [get] treeview lines
    If Not (m_cTreeView Is Nothing) Then
        TreeViewHasLines = m_cTreeView.HasLines
    End If
End Property

Public Property Let TreeViewHasLines(ByVal PropVal As Boolean)
'/* [let] treeview lines
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.HasLines = PropVal
    End If
End Property

Public Property Get TreeViewHasRootLines() As Boolean
Attribute TreeViewHasRootLines.VB_MemberFlags = "40"
'/* [get] treeview root lines
    If Not (m_cTreeView Is Nothing) Then
        TreeViewHasRootLines = m_cTreeView.HasRootLines
    End If
End Property

Public Property Let TreeViewHasRootLines(ByVal PropVal As Boolean)
'/* [let] treeview lines
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.HasRootLines = PropVal
    End If
End Property

Public Property Get TreeViewHeight() As Long
Attribute TreeViewHeight.VB_MemberFlags = "40"
'/* [get] treeview height
    If Not (m_cTreeView Is Nothing) Then
        TreeViewHeight = m_cTreeView.Height
    End If
End Property

Public Property Let TreeViewHeight(ByVal PropVal As Long)
'/* [let] treeview height
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.Height = PropVal
    End If
End Property

Public Property Get TreeViewItemHeight() As Long
Attribute TreeViewItemHeight.VB_MemberFlags = "40"
'/* [get] treeview item height
    If Not (m_cTreeView Is Nothing) Then
        TreeViewItemHeight = m_cTreeView.ItemHeight
    End If
End Property

Public Property Let TreeViewItemHeight(ByVal PropVal As Long)
'/* [let] treeview item height
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.ItemHeight = PropVal
    End If
End Property

Public Property Get TreeViewItemIndent() As Long
Attribute TreeViewItemIndent.VB_MemberFlags = "40"
'/* [get] treeview item indent
    If Not (m_cTreeView Is Nothing) Then
        TreeViewItemIndent = m_cTreeView.ItemIndent
    End If
End Property

Public Property Let TreeViewItemIndent(ByVal PropVal As Long)
'/* [let] treeview item indent
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.ItemIndent = PropVal
    End If
End Property

Public Property Get TreeViewLabelEdit() As Boolean
Attribute TreeViewLabelEdit.VB_MemberFlags = "40"
'/* [get] treeview item edit
    If Not (m_cTreeView Is Nothing) Then
        TreeViewLabelEdit = m_cTreeView.LabelEdit
    End If
End Property

Public Property Let TreeViewLabelEdit(ByVal PropVal As Boolean)
'/* [let] treeview item edit
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.LabelEdit = PropVal
    End If
End Property

Public Property Get TreeViewLineColor() As Long
Attribute TreeViewLineColor.VB_MemberFlags = "40"
'/* [get] treeview item edit
    If Not (m_cTreeView Is Nothing) Then
        TreeViewLineColor = m_cTreeView.LineColor
    End If
End Property

Public Property Let TreeViewLineColor(ByVal PropVal As Long)
'/* [get] treeview line color
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.LineColor = PropVal
    End If
End Property

Public Property Get TreeViewNodeGhosted(ByVal lNode As Long) As Boolean
Attribute TreeViewNodeGhosted.VB_MemberFlags = "40"
'/* [get] node ghosted
    If Not (m_cTreeView Is Nothing) Then
        TreeViewNodeGhosted = m_cTreeView.NodeGhosted(lNode)
    End If
End Property

Public Property Let TreeViewNodeGhosted(ByVal lNode As Long, _
                                        ByVal PropVal As Boolean)
'/* [let] node ghosted
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.NodeGhosted(lNode) = PropVal
    End If
End Property

Public Property Get TreeViewNodeHilited(ByVal lNode As Long) As Boolean
Attribute TreeViewNodeHilited.VB_MemberFlags = "40"
'/* [get] node hilted
    If Not (m_cTreeView Is Nothing) Then
        TreeViewNodeHilited = m_cTreeView.NodeHilited(lNode)
    End If
End Property

Public Property Let TreeViewNodeHilited(ByVal lNode As Long, _
                                        ByVal PropVal As Boolean)
'/* [let] node hilted
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.NodeHilited(lNode) = PropVal
    End If
End Property

Public Property Get TreeViewNodeChecked(ByVal lNode As Long) As Boolean
Attribute TreeViewNodeChecked.VB_MemberFlags = "40"
'/* [get] node checked
    If Not (m_cTreeView Is Nothing) Then
        TreeViewNodeChecked = m_cTreeView.NodeChecked(lNode)
    End If
End Property

Public Property Let TreeViewNodeChecked(ByVal lNode As Long, _
                                        ByVal PropVal As Boolean)
'/* [let] node checked
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.NodeChecked(lNode) = PropVal
    End If
End Property

Public Property Get TreeViewNodeCount() As Long
Attribute TreeViewNodeCount.VB_MemberFlags = "40"
'/* [get] node count
    If Not (m_cTreeView Is Nothing) Then
        TreeViewNodeCount = m_cTreeView.NodeCount
    End If
End Property

Public Property Get TreeViewNodeExpanded(ByVal lNode As Long) As Boolean
Attribute TreeViewNodeExpanded.VB_MemberFlags = "40"
'/* [get] node expanded
    If Not (m_cTreeView Is Nothing) Then
        TreeViewNodeExpanded = m_cTreeView.NodeExpanded(lNode)
    End If
End Property

Public Property Get TreeViewNodeParent(ByVal lNode As Long) As Boolean
Attribute TreeViewNodeParent.VB_MemberFlags = "40"
'/* [get] node parent
    If Not (m_cTreeView Is Nothing) Then
        TreeViewNodeParent = m_cTreeView.NodeParent(lNode)
    End If
End Property

Public Property Get TreeViewNodePlusMinusButton(ByVal lNode As Long) As Boolean
Attribute TreeViewNodePlusMinusButton.VB_MemberFlags = "40"
'/* [get] node glyph
    If Not (m_cTreeView Is Nothing) Then
        TreeViewLabelEdit = m_cTreeView.NodePlusMinusButton(lNode)
    End If
End Property

Public Property Let TreeViewNodePlusMinusButton(ByVal lNode As Long, _
                                                ByVal PropVal As Boolean)
'/* [let] node glyph
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.NodePlusMinusButton(lNode) = PropVal
    End If
End Property

Public Property Get TreeViewNodePrevious(ByVal lNode As Long) As Long
Attribute TreeViewNodePrevious.VB_MemberFlags = "40"
'/* [get] previous node
    If Not (m_cTreeView Is Nothing) Then
        TreeViewNodePrevious = m_cTreeView.NodePrevious(lNode)
    End If
End Property

Public Property Get TreeViewNodeFirstSibling(ByVal lNode As Long) As Long
Attribute TreeViewNodeFirstSibling.VB_MemberFlags = "40"
'/* [get] node first sibling
    If Not (m_cTreeView Is Nothing) Then
        TreeViewNodeFirstSibling = m_cTreeView.NodeFirstSibling(lNode)
    End If
End Property

Public Property Get TreeViewNodeFirstVisible() As Long
Attribute TreeViewNodeFirstVisible.VB_MemberFlags = "40"
'/* [get] first visible node
    If Not (m_cTreeView Is Nothing) Then
        TreeViewNodeFirstVisible = m_cTreeView.NodeFirstVisible
    End If
End Property

Public Property Get TreeViewNodeImage(ByVal lNode As Long) As Long
Attribute TreeViewNodeImage.VB_MemberFlags = "40"
'/* [get] node image
    If Not (m_cTreeView Is Nothing) Then
        TreeViewNodeImage = m_cTreeView.NodeImage(lNode)
    End If
End Property

Public Property Let TreeViewNodeImage(ByVal lNode As Long, _
                                      ByVal PropVal As Long)
'/* [let] node image
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.NodeImage(lNode) = PropVal
    End If
End Property

Public Property Get TreeViewNodeLastVisible() As Long
Attribute TreeViewNodeLastVisible.VB_MemberFlags = "40"
'/* [get] last visible node
    If Not (m_cTreeView Is Nothing) Then
        TreeViewNodeLastVisible = m_cTreeView.NodeLastVisible
    End If
End Property

Public Property Get TreeViewNodeRoot() As Long
Attribute TreeViewNodeRoot.VB_MemberFlags = "40"
'/* [get] root node
    If Not (m_cTreeView Is Nothing) Then
        TreeViewNodeRoot = m_cTreeView.NodeRoot
    End If
End Property

Public Property Get TreeViewNodeSelectedImage(ByVal lNode As Long) As Long
Attribute TreeViewNodeSelectedImage.VB_MemberFlags = "40"
'/* [get] node selected image
    If Not (m_cTreeView Is Nothing) Then
        TreeViewNodeSelectedImage = m_cTreeView.NodeSelectedImage(lNode)
    End If
End Property

Public Property Let TreeViewNodeSelectedImage(ByVal lNode As Long, _
                                              ByVal PropVal As Long)
'/* [let] node selected image
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.NodeSelectedImage(lNode) = PropVal
    End If
End Property

Public Property Get TreeViewNodeTag(ByVal lNode As Long) As String
Attribute TreeViewNodeTag.VB_MemberFlags = "40"
'/* [get] node tag
    If Not (m_cTreeView Is Nothing) Then
        TreeViewNodeTag = m_cTreeView.NodeTag(lNode)
    End If
End Property

Public Property Let TreeViewNodeTag(ByVal lNode As Long, _
                                    ByVal PropVal As String)
'/* [let] node tag
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.NodeTag(lNode) = PropVal
    End If
End Property

Public Property Get TreeViewNodeText(ByVal lNode As Long) As String
Attribute TreeViewNodeText.VB_MemberFlags = "40"
'/* [get] node text
    If Not (m_cTreeView Is Nothing) Then
        TreeViewNodeText = m_cTreeView.NodeText(lNode)
    End If
End Property

Public Property Let TreeViewNodeText(ByVal lNode As Long, _
                                     ByVal PropVal As String)
'/* [let] node text
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.NodeText(lNode) = PropVal
    End If
End Property

Public Property Get TreeViewOLEDragMode() As ETVOLEDragConstants
Attribute TreeViewOLEDragMode.VB_MemberFlags = "40"
'/* [get] drag mode
    If Not (m_cTreeView Is Nothing) Then
        TreeViewOLEDragMode = m_cTreeView.OLEDragMode
    End If
End Property

Public Property Let TreeViewOLEDragMode(ByVal PropVal As ETVOLEDragConstants)
'/* [let] drag mode
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.OLEDragMode = PropVal
    End If
End Property

Public Property Get TreeViewScrollBarAlignment() As ESCScrollBarAlignment
Attribute TreeViewScrollBarAlignment.VB_MemberFlags = "40"
'/* [get] tv scrollbar alignment
    TreeViewScrollBarAlignment = m_eTvScrollBarAlignment
End Property

Public Property Let TreeViewScrollBarAlignment(ByVal PropVal As ESCScrollBarAlignment)
'/* [let] tv scrollbar alignment
    If Not (m_cTreeView Is Nothing) Then
        If Not (m_lTVHwnd = 0) Then
            If (PropVal = escRightAlign) Then
                WindowStyle m_lTVHwnd, GWL_EXSTYLE, 0, WS_EX_LEFTSCROLLBAR
            Else
                WindowStyle m_lTVHwnd, GWL_EXSTYLE, WS_EX_LEFTSCROLLBAR, 0
            End If
            Resize
        End If
    End If
    m_eTvScrollBarAlignment = PropVal
End Property

Public Property Get TreeViewSelectedBackColor() As Long
Attribute TreeViewSelectedBackColor.VB_MemberFlags = "40"
'/* [get] treeview selected backcolor
    If Not (m_cTreeView Is Nothing) Then
        TreeViewSelectedBackColor = m_cTreeView.SelectedBackColor
    End If
End Property

Public Property Let TreeViewSelectedBackColor(ByVal PropVal As Long)
'/* [let] treeview selected backcolor
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.SelectedBackColor = PropVal
    End If
End Property

Public Property Get TreeViewSizeable() As Boolean
Attribute TreeViewSizeable.VB_MemberFlags = "40"
'/* [get] treeview can size
    TreeViewSizeable = TreeViewSizeable
End Property

Public Property Let TreeViewSizeable(ByVal PropVal As Boolean)
'/* [let] treeview can size
    m_bTreeViewSizeable = PropVal
End Property

Public Property Get TreeViewUseUnicode() As Boolean
Attribute TreeViewUseUnicode.VB_MemberFlags = "40"
'/* [get] tree unicode state
    If Not (m_cTreeView Is Nothing) Then
        TreeViewUseUnicode = m_cTreeView.UseUnicode
    End If
End Property

Public Property Let TreeViewUseUnicode(ByVal PropVal As Boolean)
'/* [let] tree unicode state
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.UseUnicode = PropVal
    End If
End Property

Public Property Get TreeViewWidth() As Long
Attribute TreeViewWidth.VB_MemberFlags = "40"
'/* [get] treeview width
    If Not (m_cTreeView Is Nothing) Then
        TreeViewWidth = m_cTreeView.Width
    End If
End Property

Public Property Let TreeViewWidth(ByVal PropVal As Long)
'/* [let] treeview width
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.Width = PropVal
    End If
End Property


'**********************************************************************
'*                              SKINNING
'**********************************************************************

Public Property Get AdvancedEdit() As Boolean
Attribute AdvancedEdit.VB_MemberFlags = "40"
'/* [get] advanced edit state
    AdvancedEdit = m_bAdvancedEdit
End Property

Public Property Let AdvancedEdit(ByVal PropVal As Boolean)
'/* [let] enable advanced edit
    m_bAdvancedEdit = PropVal
End Property

Public Property Get AdvancedEditThemeColor() As Long
Attribute AdvancedEditThemeColor.VB_MemberFlags = "40"
'/* [get] get advanced edit theme color
    AdvancedEditThemeColor = m_lAdvancedEditThemeColor
End Property

Public Property Let AdvancedEditThemeColor(ByVal PropVal As Long)
'/* [let] set advanced edit theme color
    m_lAdvancedEditThemeColor = PropVal
End Property

Public Property Get AdvancedEditOffsetColor() As Long
Attribute AdvancedEditOffsetColor.VB_MemberFlags = "40"
'/* [get] get advanced edit theme offset color
    AdvancedEditOffsetColor = m_lAdvancedEditOffsetColor
End Property

Public Property Let AdvancedEditOffsetColor(ByVal PropVal As Long)
'/* [let] set advanced edit theme offset color
    m_lAdvancedEditOffsetColor = PropVal
End Property

Public Property Get AdvancedEditThemStyle() As EVSThemeStyle
Attribute AdvancedEditThemStyle.VB_MemberFlags = "40"
'/* [get] get advanced edit style
    AdvancedEditThemStyle = m_eEditorThemeStyle
End Property

Public Property Let AdvancedEditThemStyle(ByVal PropVal As EVSThemeStyle)
'/* [let] set advanced edit style
    m_eEditorThemeStyle = PropVal
End Property

Public Property Get AlphaBarActive() As Boolean
'/* [get] alpha bar loaded state
    If m_bAlphaIsLoaded Then
        AlphaBarActive = m_bAlphaSelectorBar
    End If
End Property

Public Property Let AlphaBarActive(ByVal PropVal As Boolean)
'/* [let] alpha bar loaded state
    If PropVal Then
        If AlphaBarTransparency = 0 Then
            AlphaBarTransparency = 100
        End If
        AlphaSelectorBar AlphaBarTransparency, m_bAlphaBarTheme
    End If
    m_bAlphaSelectorBar = PropVal
    PropertyChanged "AlphaBarActive"
End Property

Public Property Get AlphaBarTheme() As Boolean
Attribute AlphaBarTheme.VB_MemberFlags = "40"
'/* [get] use theme color
    AlphaBarTheme = m_bAlphaBarTheme
End Property

Public Property Let AlphaBarTheme(ByVal PropVal As Boolean)
'/* [let] use theme color
    m_bAlphaBarTheme = PropVal
    PropertyChanged "AlphaBarTheme"
End Property

Public Property Get AlphaBarTransparency() As Byte
'/* [get] alpha bar transparency index
    AlphaBarTransparency = m_bteAlphaTransparency
End Property

Public Property Let AlphaBarTransparency(ByVal PropVal As Byte)
'/* [let] alpha bar transparency index
    If PropVal < 70 Then
        m_bteAlphaTransparency = 70
    ElseIf PropVal > 240 Then
        m_bteAlphaTransparency = 200
    Else
        m_bteAlphaTransparency = PropVal
    End If
    PropertyChanged "AlphaBarTransparency"
End Property

Public Property Get ThemeAutoXp() As Boolean
Attribute ThemeAutoXp.VB_MemberFlags = "40"
'/* [get] auto assign xp theme
    ThemeAutoXp = m_bThemeAutoXp
End Property

Public Property Let ThemeAutoXp(ByVal PropVal As Boolean)
'/* [let] auto assign xp theme
    If m_bIsXp Then
        If PropVal Then
            Select Case GetThemeName
            Case HomeStead
                m_eHeaderSkinStyle = evsXpGreen
                m_eScrollBarSkinStyle = evsXpGreen
                m_eCheckBoxSkinStyle = evsXpGreen
            Case Metallic
                m_eHeaderSkinStyle = evsXpSilver
                m_eScrollBarSkinStyle = evsXpSilver
                m_eCheckBoxSkinStyle = evsXpSilver
            Case NormalColor
                m_eHeaderSkinStyle = evsXpBlue
                m_eScrollBarSkinStyle = evsXpBlue
                m_eCheckBoxSkinStyle = evsXpBlue
            Case Else
                m_eHeaderSkinStyle = evsXpBlue
                m_eScrollBarSkinStyle = evsXpBlue
                m_eCheckBoxSkinStyle = evsXpBlue
            End Select
            m_bThemeAutoXp = PropVal
            If Not (m_lHGHwnd = 0) Then
                SkinScrollBars m_eScrollBarSkinStyle, False
                SkinHeaders m_eHeaderSkinStyle, m_oHdrForeClr, m_oHdrHighLiteClr, m_oHdrPressedClr, False
                If m_bCheckBoxes Then
                    SkinCheckBox m_eCheckBoxSkinStyle, False
                End If
                If m_bHasTreeView Then
                    m_cTreeView.SkinCheckBox m_eCheckBoxSkinStyle
                    m_cTreeView.SkinScrollBars m_eScrollBarSkinStyle
                End If
                Resize
            End If
        End If
    End If
End Property

Private Function AlphaBarReset() As Boolean
'/* reset alpha bar

    If Not (m_cSelectorBar Is Nothing) Then
        Set m_cSelectorBar = Nothing
        Set ISelectorBar = Nothing
    End If

End Function

Public Sub AlphaSelectorBar(ByVal btTransparency As Byte, _
                            ByVal bUseThemeColor As Boolean)

'/* create alpha selector bar

    AlphaBarTransparency = btTransparency
    m_bAlphaBarTheme = bUseThemeColor
    '/* reset
    If Not (m_cSelectorBar Is Nothing) Then
        AlphaBarReset
    End If
    '/* load image
    Set ISelectorBar = LoadResPicture("SELECTORBAR", vbResBitmap)
    '/* create dc
    Set m_cSelectorBar = New clsStoreDc
    '/* alpha and colorization
    With m_cSelectorBar
        .UseAlpha = True
        .CreateFromPicture ISelectorBar
        If bUseThemeColor Then
            .ColorizeImage m_oThemeColor, m_sngLuminence
        End If
    End With
    m_bAlphaIsLoaded = True
    
End Sub

Private Function LoadCheckBoxImages() As Boolean
'/* load checkbox skin images

    ResetSkinnedCheckboxes

    Select Case m_eCheckBoxSkinStyle
    '/* azure
    Case 0
        Set m_IChecked = LoadResPicture("AZURE-CHKPUSHED", vbResBitmap)
        Set m_IUnChecked = LoadResPicture("AZURE-CHKEMPTY", vbResBitmap)
        Set m_IChkDisabled = LoadResPicture("AZURE-CHKDISABLED", vbResBitmap)
    '/* classic
    Case 1
        Set m_IChecked = LoadResPicture("CLASSIC-CHKPUSHED", vbResBitmap)
        Set m_IUnChecked = LoadResPicture("CLASSIC-CHKEMPTY", vbResBitmap)
        Set m_IChkDisabled = LoadResPicture("CLASSIC-CHKDISABLED", vbResBitmap)
    '/* gloss
    Case 2
        Set m_IChecked = LoadResPicture("GLOSS-CHKPUSHED", vbResBitmap)
        Set m_IUnChecked = LoadResPicture("GLOSS-CHKEMPTY", vbResBitmap)
        Set m_IChkDisabled = LoadResPicture("GLOSS-CHKDISABLED", vbResBitmap)
    '/* metal
    Case 3
        Set m_IChecked = LoadResPicture("METAL-CHKPUSHED", vbResBitmap)
        Set m_IUnChecked = LoadResPicture("METAL-CHKEMPTY", vbResBitmap)
        Set m_IChkDisabled = LoadResPicture("METAL-CHKDISABLED", vbResBitmap)
    '/* xp
    Case 4, 5, 6
        Set m_IChecked = LoadResPicture("XP-CHKPUSHED", vbResBitmap)
        Set m_IUnChecked = LoadResPicture("XP-CHKEMPTY", vbResBitmap)
        Set m_IChkDisabled = LoadResPicture("XP-CHKDISABLED", vbResBitmap)
    '/* vista
    Case 7
        Set m_IChecked = LoadResPicture("VISTA-CHKPUSHED", vbResBitmap)
        Set m_IUnChecked = LoadResPicture("VISTA-CHKEMPTY", vbResBitmap)
        Set m_IChkDisabled = LoadResPicture("VISTA-CHKDISABLED", vbResBitmap)
    End Select

    '/* success
    LoadCheckBoxImages = True

End Function

Private Sub ResetSkinnedCheckboxes()
'/* reset checkbox images

    ImlClear m_lImlStateHndl
    If Not m_IChecked Is Nothing Then Set m_IChecked = Nothing
    If Not m_cChkCheckDc Is Nothing Then Set m_cChkCheckDc = Nothing
    If Not m_IUnChecked Is Nothing Then Set m_IUnChecked = Nothing
    If Not m_cChkUnCheckDc Is Nothing Then Set m_cChkUnCheckDc = Nothing
    If Not m_IChkDisabled Is Nothing Then Set m_IChkDisabled = Nothing
    If Not m_cChkDisableDc Is Nothing Then Set m_cChkDisableDc = Nothing

End Sub

Public Property Get ImlUseAlphaIcons() As Boolean
    ImlUseAlphaIcons = m_bImlUseAlphaIcons
End Property

Public Property Let ImlUseAlphaIcons(ByVal PropVal As Boolean)
    m_cHeaderIcons.UseGdiPlus = PropVal
    m_cCellIcons.UseGdiPlus = PropVal
    If m_bHasInitialized Then
        m_cCellTips.UseGdiPlus = PropVal
    End If
    If m_bHasTreeView Then
        m_cTreeIcons.UseGdiPlus = PropVal
    End If
    m_bImlUseAlphaIcons = PropVal
End Property

Public Sub ImlRemoveIcon(ByVal lImlHnd As Long, _
                         ByVal lIndex As Long)

    If Not (lImlHnd = 0) Then
        ImageList_Remove lImlHnd, lIndex
    End If

End Sub

Public Sub ImlClear(ByVal lImlHnd As Long)

Dim lCt     As Long
Dim lCount  As Long

    If Not (lImlHnd = 0) Then
        lCount = ImageList_GetImageCount(lImlHnd) - 1
        For lCt = lCount To 0 Step -1
            ImlRemoveIcon lImlHnd, lCt
        Next lCt
    End If

End Sub

Public Sub SkinCheckBox(ByVal eCheckBoxStyle As EVSThemeStyle, _
                        ByVal bUseThemeColors As Boolean)

'/* skin the listview checkboxes

Dim lMask As Long

    m_eCheckBoxSkinStyle = eCheckBoxStyle
    m_bUseCheckBoxTheme = bUseThemeColors
    '/* system image sizes
    CheckBoxMetrics

    '/* load images
    LoadCheckBoxImages
    '/* image dc's
    Set m_cChkCheckDc = New clsStoreDc
    Set m_cChkUnCheckDc = New clsStoreDc
    Set m_cChkDisableDc = New clsStoreDc
    '/* create dc's
    If m_bUseCheckBoxTheme Then
        With m_cChkUnCheckDc
            '/* create image dc
            .CreateFromPicture m_IUnChecked
            '/* colorize
            .ColorizeImage m_oThemeColor, m_sngLuminence
            '/* new mask color
            lMask = GetMaskColor(.hdc)
            '/* extract bitmap handle
            ImlStateAddBmp .ExtractBitmap, lMask
        End With
        With m_cChkCheckDc
            .CreateFromPicture m_IChecked
            .ColorizeImage m_oThemeColor, m_sngLuminence
            ImlStateAddBmp .ExtractBitmap, lMask
        End With
        With m_cChkDisableDc
            .CreateFromPicture m_IChkDisabled
            .ColorizeImage m_oThemeColor, m_sngLuminence
            ImlStateAddBmp .ExtractBitmap, lMask
        End With
        Set m_cChkCheckDc = Nothing
        Set m_cChkUnCheckDc = Nothing
        Set m_cChkDisableDc = Nothing
    Else
        ImlStateAddBmp m_IUnChecked.handle, &HFF00FF
        ImlStateAddBmp m_IChecked.handle, &HFF00FF
        ImlStateAddBmp m_IChkDisabled.handle, &HFF00FF
    End If

End Sub

Public Sub SkinHeaders(ByVal eSkinStyle As EVSThemeStyle, _
                       ByVal oFontForecolor As OLE_COLOR, _
                       ByVal oFontHighliteColor As OLE_COLOR, _
                       ByVal oFontPressedColor As OLE_COLOR, _
                       ByVal bUseThemeColors As Boolean)

'/* use header skin

    '/* skin params
    If Not (m_cSkinHeader Is Nothing) Then
        If m_bSkinHeader Then
            m_cSkinHeader.ResetHeaderSkin
        End If
        '/* load properties
        m_eHeaderSkinStyle = eSkinStyle
        m_oHdrForeClr = oFontForecolor
        m_oHdrHighLiteClr = oFontHighliteColor
        m_oHdrPressedClr = oFontPressedColor
        m_bUseThemeColors = bUseThemeColors
        '/* pass through to skinheaderclass
        With m_cSkinHeader
            .HeaderForeColor = m_oHdrForeClr
            .HeaderHighLite = m_oHdrHighLiteClr
            .HeaderPressed = m_oHdrPressedClr
            .HeaderIml = m_cHeaderIcons.hIml
            .HeaderLuminence = m_eThemeLuminence
            .HeaderSkinStyle = m_eHeaderSkinStyle
            .HeaderThemeColor = m_oThemeColor
            .UseUnicode = m_bUseUnicode
            .HeaderTextEffect = m_lHeaderTextEffect
            .SetFont m_oHeaderFont
            .ColumnVerticalText = m_bColumnVerticalText
            .UseHeaderTheme = m_bUseThemeColors
            .FontRightLeading = m_bFontRightLeading
            .ToolTips = True
            .LoadSkin
            m_bSkinHeader = True
        End With
    End If

End Sub

Public Sub SkinScrollBars(ByVal eSkinStyle As EVSThemeStyle, _
                          ByVal bUseThemeColors As Boolean)

'/* skin scrollbars

    '/* reset dc class
    If Not (m_cSkinScrollBars Is Nothing) Then
        m_cSkinScrollBars.ResetScrollBarSkin
    Else
        Set m_cSkinScrollBars = New clsSkinScrollbars
    End If
    '/* set properties
    m_eScrollBarSkinStyle = eSkinStyle
    '/* pass though to skinscrollbars class
    With m_cSkinScrollBars
        .ScrollBarSkinStyle = m_eScrollBarSkinStyle
        .ScrollLuminence = m_eThemeLuminence
        .ScrollThemeColor = m_oThemeColor
        .SkinScrollBar = True
        .UseScrollBarTheme = bUseThemeColors
        .LoadSkin m_lHGHwnd, m_lParentHwnd
    End With
    m_bSkinScrollBars = True
    Resize

End Sub

Public Sub ThemeManager(ByVal eSkinStyle As EVSThemeStyle, _
                        Optional ByVal bUseSkinTheme As Boolean = False, _
                        Optional ByVal lSkinThemeColor As Long = -1, _
                        Optional ByVal eThemeLuminence As ESTThemeLuminence = estThemeSoft, _
                        Optional ByVal lColumnFontColor As Long = &H404040, _
                        Optional ByVal lColumnFontHiliteColor As Long = &H808080, _
                        Optional ByVal lColumnFontPressedColor As Long = &H0, _
                        Optional ByVal lOptionFormBackColor As Long = -1, _
                        Optional ByVal lOptionFormOffsetColor As Long = -1, _
                        Optional ByVal lOptionFormForeColor As Long = &H0, _
                        Optional ByVal bteOptionFormTransparency As Byte = 0, _
                        Optional ByVal bOptionFormGradient As Boolean = False, _
                        Optional ByVal bUseXPColors As Boolean = False, _
                        Optional ByVal lOptionControlColor As Long = -1, _
                        Optional ByVal lOptionControlForeColor As Long = &H0, _
                        Optional ByVal bIncludeAdvancedEdit As Boolean = False, _
                        Optional ByVal bIncludeColumnTip As Boolean = False, _
                        Optional ByVal bIncludeFilter As Boolean = False, _
                        Optional ByVal bIncludeCellTip As Boolean = False, _
                        Optional ByVal bIncludeTreeView As Boolean = False)

'/* centralized style management

    '/* defaults
    If (lOptionFormBackColor = -1) Then
        lOptionFormBackColor = GetSysColor(vbButtonFace And &H1F)
    End If
    If (lOptionControlColor = -1) Then
        lOptionControlColor = lOptionFormBackColor
    End If
    
    '/* skin theme colors
    If bUseSkinTheme Then
        If lSkinThemeColor = -1 Then
            bUseSkinTheme = False
        Else
            m_bUseThemeColors = bUseSkinTheme
            If m_bXPColors Then
                lSkinThemeColor = m_cRender.XPShift(lSkinThemeColor)
            End If
            m_oThemeColor = lSkinThemeColor
            m_eThemeLuminence = eThemeLuminence
        End If
    End If
    
    '/* column filter box
    If bIncludeFilter Then
        If Not (lOptionFormBackColor = -1) Then
            m_oFilterBackColor = lOptionFormBackColor
            If Not (lOptionFormOffsetColor = -1) Then
                m_oFilterOffsetColor = lOptionFormOffsetColor
            Else
                m_bFilterGradient = False
            End If
        End If
        m_oFilterControlColor = lOptionControlColor
        m_oFilterControlForeColor = lOptionControlForeColor
        m_oFilterTitleColor = lOptionFormForeColor
        m_oFilterForeColor = lOptionFormForeColor
        m_lFilterTransparency = bteOptionFormTransparency
        m_bFilterGradient = bOptionFormGradient
        m_bFilterXPColors = bUseXPColors
    End If
        
    '/* cell tool tips
    If bIncludeCellTip Then
        If Not (lOptionFormBackColor = -1) Then
            m_oCellTipColor = lOptionFormBackColor
            If Not (lOptionFormOffsetColor = -1) Then
                m_oCellTipOffsetColor = lOptionFormOffsetColor
            Else
                m_bCellTipGradient = False
            End If
        End If
        m_oCellTipForeColor = lOptionControlForeColor
        m_lCellTipTransparency = bteOptionFormTransparency
        m_bCellTipGradient = bOptionFormGradient
        m_bCellTipXPColors = bUseXPColors
    End If
    
    '/* column tooltip
    If bIncludeColumnTip Then
        If Not (lOptionFormBackColor = -1) Then
            ColumnTipColor = lOptionFormBackColor
            If Not (lOptionFormOffsetColor = -1) Then
                ColumnTipOffsetColor = lOptionFormOffsetColor
            Else
                ColumnTipGradient = False
            End If
        End If
        ColumnTipForeColor = lOptionControlForeColor
        ColumnTipTransparency = bteOptionFormTransparency
        ColumnTipGradient = bOptionFormGradient
        ColumnTipXPColors = bUseXPColors
    End If
    
    '/* advanced edit option
    If bIncludeAdvancedEdit Then
        m_eEditorThemeStyle = eSkinStyle
        m_lAdvancedEditThemeColor = lOptionControlColor
        m_lAdvancedEditOffsetColor = lColumnFontHiliteColor
    End If
    
    '/* treeview styles
    If bIncludeTreeView Then
        If Not (m_cTreeView Is Nothing) Then
            With m_cTreeView
                If (lOptionFormBackColor = -1) Then
                    If m_bXPColors Then
                        .BackColor = m_cRender.XPShift(m_oBackColor)
                    Else
                        .BackColor = m_oBackColor
                    End If
                Else
                    If m_bXPColors Then
                        .BackColor = m_cRender.XPShift(lOptionFormBackColor)
                    Else
                        .BackColor = m_oBackColor
                    End If
                End If
                .ForeColor = m_oForeColor
                If bUseSkinTheme Then
                    .SkinCheckBox eSkinStyle, lSkinThemeColor, m_sngLuminence
                    .SkinScrollBars eSkinStyle, lSkinThemeColor, m_sngLuminence
                Else
                    .SkinCheckBox eSkinStyle
                    .SkinScrollBars eSkinStyle
                End If
            End With
        End If
    End If
    
    Select Case eSkinStyle
    Case evsAzure
        SkinScrollBars evsAzure, bUseSkinTheme
        SkinHeaders evsAzure, lColumnFontColor, lColumnFontHiliteColor, lColumnFontPressedColor, bUseSkinTheme
        If m_bCheckBoxes Then
            SkinCheckBox evsAzure, bUseSkinTheme
        End If
        
    Case evsClassic
        SkinScrollBars evsClassic, bUseSkinTheme
        SkinHeaders evsClassic, lColumnFontColor, lColumnFontHiliteColor, lColumnFontPressedColor, bUseSkinTheme
        If m_bCheckBoxes Then
            SkinCheckBox evsClassic, bUseSkinTheme
        End If
        
    Case evsGloss
        SkinScrollBars evsGloss, bUseSkinTheme
        SkinHeaders evsGloss, lColumnFontColor, lColumnFontHiliteColor, lColumnFontPressedColor, bUseSkinTheme
        If m_bCheckBoxes Then
            SkinCheckBox evsGloss, bUseSkinTheme
        End If
        
    Case evsMetallic
        SkinScrollBars evsMetallic, bUseSkinTheme
        SkinHeaders evsMetallic, lColumnFontColor, lColumnFontHiliteColor, lColumnFontPressedColor, bUseSkinTheme
        If m_bCheckBoxes Then
            SkinCheckBox evsMetallic, bUseSkinTheme
        End If
        
    Case evsXpSilver
        SkinScrollBars evsXpSilver, bUseSkinTheme
        SkinHeaders evsXpSilver, lColumnFontColor, lColumnFontHiliteColor, lColumnFontPressedColor, bUseSkinTheme
        If m_bCheckBoxes Then
            SkinCheckBox evsXpSilver, bUseSkinTheme
        End If
    
    Case evsXpBlue
        SkinScrollBars evsXpBlue, bUseSkinTheme
        SkinHeaders evsXpBlue, lColumnFontColor, lColumnFontHiliteColor, lColumnFontPressedColor, bUseSkinTheme
        If m_bCheckBoxes Then
            SkinCheckBox evsXpBlue, bUseSkinTheme
        End If
        
    Case evsXpGreen
        SkinScrollBars evsXpGreen, bUseSkinTheme
        SkinHeaders evsXpGreen, lColumnFontColor, lColumnFontHiliteColor, lColumnFontPressedColor, bUseSkinTheme
        If m_bCheckBoxes Then
            SkinCheckBox evsXpGreen, bUseSkinTheme
        End If
        
    Case evsVistaArrow
        SkinScrollBars evsVistaArrow, bUseSkinTheme
        SkinHeaders evsVistaArrow, lColumnFontColor, lColumnFontHiliteColor, lColumnFontPressedColor, bUseSkinTheme
        If m_bCheckBoxes Then
            SkinCheckBox evsVistaArrow, bUseSkinTheme
        End If
    Case evsSilver
        SkinScrollBars evsSilver, bUseSkinTheme
        SkinHeaders evsSilver, lColumnFontColor, lColumnFontHiliteColor, lColumnFontPressedColor, bUseSkinTheme
        If m_bCheckBoxes Then
            SkinCheckBox evsMetallic, bUseSkinTheme
        End If
    End Select

End Sub

Public Sub UnSkinAll()
'/* remove all skinning

    UnSkinHeaders
    UnSkinScrollBars

End Sub

Public Sub UnSkinHeaders()
'/* reset skinned header
    Set m_cSkinHeader = Nothing
End Sub

Public Sub UnSkinScrollBars()
'/* unskin scrollbars

    If Not (m_cSkinScrollBars Is Nothing) Then
        Set m_cSkinScrollBars = Nothing
    End If
    m_bSkinScrollBars = False
    
End Sub


'**********************************************************************
'*                            GRID PROPERTIES
'**********************************************************************


Public Property Get BackColor() As OLE_COLOR
'*/ retrieve list backcolor
    BackColor = m_oBackColor
End Property

Public Property Let BackColor(ByVal PropVal As OLE_COLOR)
'*/ change list backcolor
    OleTranslateColor PropVal, 0&, m_oBackColor
    If m_bXPColors Then
        m_oBackColor = m_cRender.XPShift(PropVal)
    Else
        m_oBackColor = PropVal
    End If
    If Not (m_lHGHwnd = 0) Then
        If m_bEnabled Then
            SendMessageLongA m_lHGHwnd, LVM_SETBKCOLOR, 0&, m_oBackColor
        End If
    End If
    GridRefresh False
    PropertyChanged "BackColor"
End Property

Public Property Get BorderStyle() As EBSBorderStyle
'*/ [get] list borderstyle
    BorderStyle = m_eBorderStyle
End Property

Public Property Let BorderStyle(ByVal PropVal As EBSBorderStyle)
'*/ [let] change list borderstyle
    SetBorderStyle UserControl.hwnd, PropVal
    m_eBorderStyle = PropVal
    UserControl_Resize
    PropertyChanged "BorderStyle"
End Property

Public Property Get CellUseDecoration() As Boolean
Attribute CellUseDecoration.VB_MemberFlags = "40"
'*/ [get] get structured cell decoration
    CellUseDecoration = m_bCellDecoration
End Property

Public Property Let CellUseDecoration(ByVal PropVal As Boolean)
'*/ [let] enable structured cell decoration
    m_bCellDecoration = PropVal
End Property

Public Property Get CellEdit() As Boolean
'*/ [get] enable cell editing
    CellEdit = m_bCellEdit
End Property

Public Property Let CellEdit(ByVal PropVal As Boolean)
'*/ [let] enable cell editing
    m_bCellEdit = PropVal
    EditUpdateLabel
    PropertyChanged "CellEdit"
End Property

Public Property Get CellFocusedColor() As OLE_COLOR
'*/ [get] cell focus color
    CellFocusedColor = m_oCellFocusedColor
End Property

Public Property Let CellFocusedColor(ByVal PropVal As OLE_COLOR)
'*/ [let] cell focus color
    m_oCellFocusedColor = PropVal
    PropertyChanged "CellFocusedColor"
End Property

Public Property Get CellHiliteColor() As Long
Attribute CellHiliteColor.VB_MemberFlags = "40"
'*/ [get] ole cell hilite color
    CellHiliteColor = m_lCellHiliteColor
End Property

Public Property Let CellHiliteColor(ByVal PropVal As Long)
'*/ [let] ole cell hilite color
    m_lCellHiliteColor = PropVal
End Property

Public Property Get CellHiliteStyle() As ECHCellHiliteStyle
Attribute CellHiliteStyle.VB_MemberFlags = "40"
'*/ [get] ole cell hilite style
    CellHiliteStyle = m_eCellHiliteStyle
End Property

Public Property Let CellHiliteStyle(ByVal PropVal As ECHCellHiliteStyle)
'*/ [let] ole cell hilite style
    m_eCellHiliteStyle = PropVal
End Property

Public Property Get CellHotTrack() As Boolean
Attribute CellHotTrack.VB_MemberFlags = "40"
'*/ [get] enable cell hot ttracking
    CellHotTrack = m_bCellHotTrack
End Property

Public Property Let CellHotTrack(ByVal PropVal As Boolean)
'*/ [let] enable cell hot ttracking
    m_bCellHotTrack = PropVal
End Property

Public Property Get CellIcon(ByVal lRow As Long, _
                             ByVal lColumn As Long) As Long
Attribute CellIcon.VB_MemberFlags = "40"
'*/ [get] return icon index
    If Not (m_lHGHwnd = 0) Then
        If Not m_bVirtualMode Then
            If (lRow < m_lRowCount) Then
                If (lColumn < m_lColumnCount) Then
                    CellIcon = m_cGridItem(lRow).Icon(lColumn)
                End If
            End If
        End If
    End If
End Property

Public Property Let CellIcon(ByVal lRow As Long, _
                             ByVal lColumn As Long, _
                             ByVal lIcon As Long)
'*/ [let] change icon index
    If Not (m_lHGHwnd = 0) Then
        If Not m_bVirtualMode Then
            If (lRow < m_lRowCount) Then
                If (lColumn < m_lColumnCount) Then
                    m_cGridItem(lRow).Icon(lColumn) = lIcon
                End If
            End If
        End If
    End If
End Property

Public Property Get CellSelectedColor() As OLE_COLOR
'*/ [get] cell selected color
    CellSelectedColor = m_oCellSelectedColor
End Property

Public Property Let CellSelectedColor(ByVal PropVal As OLE_COLOR)
'*/ [let] cell selected color
    m_oCellSelectedColor = PropVal
    PropertyChanged "CellSelectedColor"
End Property

Public Property Get CheckBoxes() As Boolean
'*/ [get] retrieve checkbox state
    CheckBoxes = m_bCheckBoxes
End Property

Public Property Let CheckBoxes(ByVal PropVal As Boolean)
'*/ [let] change checkbox state
    If Not (m_lHGHwnd = 0) Then
        If PropVal Then
            InitImlState
            SkinCheckBox m_eCheckBoxSkinStyle, m_bUseCheckBoxTheme
            SetExtendedStyle LVS_EX_CHECKBOXES, 0
        Else
            DestroyImlState
            SetExtendedStyle 0, LVS_EX_CHECKBOXES
        End If
        m_bCheckBoxes = PropVal
        '/* force row size recalculation
        RowHeightChange
    End If
    m_bCheckBoxes = PropVal
    PropertyChanged "CheckBoxes"
End Property

Public Property Get Checked(ByVal lIndex As Long) As Boolean
Attribute Checked.VB_MemberFlags = "40"
'/* [get] checkbox check state
    If ArrayCheck(m_cGridItem) Then
        Checked = m_cGridItem(lIndex).Checked
    End If
End Property

Public Property Let Checked(ByVal lIndex As Long, _
                            ByVal bChecked As Boolean)

'/* [let] checkbox check state

    If ArrayCheck(m_cGridItem) Then
        m_cGridItem(lIndex).Checked = bChecked
    End If
    GridRefresh False
    
End Property

Public Property Get ColumnAlign(ByVal lColumn As Long) As ECAColumnAlign
Attribute ColumnAlign.VB_MemberFlags = "40"
'*/ [get] retieve a columns text alignment

Dim uLVC As LVCOLUMN

    If Not (m_lHdrHwnd = 0) Then
        If (lColumn < m_lColumnCount) Then
            uLVC.Mask = LVCF_FMT
            If m_bIsNt Then
                SendMessageW m_lHGHwnd, LVM_GETCOLUMNW, lColumn, uLVC
            Else
                SendMessageA m_lHGHwnd, LVM_GETCOLUMNA, lColumn, uLVC
            End If
            ColumnAlign = (&H3 And uLVC.fmt)
        End If
    End If

End Property

Public Property Let ColumnAlign(ByVal lColumn As Long, _
                                ByVal eAlign As ECAColumnAlign)

'*/ [let] change a columns text alignment

Dim uLVC As LVCOLUMN

    If Not (m_lHdrHwnd = 0) Then
        If (lColumn < m_lColumnCount) Then
            With uLVC
                .fmt = eAlign * -(Not lColumn = 0)
                .Mask = LVCF_FMT
            End With
            If m_bIsNt Then
                SendMessageW m_lHGHwnd, LVM_SETCOLUMNW, lColumn, uLVC
            Else
                SendMessageA m_lHGHwnd, LVM_SETCOLUMNA, lColumn, uLVC
            End If
        End If
    End If

End Property

Private Property Get ColumnAtIndex(ByVal lIndex As Long) As Long
'/* [get] column array index

Dim lCt As Long

    For lCt = 0 To (m_lColumnCount - 1)
        If ColumnIndex(lCt) = lIndex Then
            ColumnAtIndex = lCt
            Exit For
        End If
    Next lCt

End Property

Public Property Get ColumnCount() As Long
Attribute ColumnCount.VB_MemberFlags = "40"
'*/ [get] retieve column count
    If Not (m_lHdrHwnd = 0) Then
        ColumnCount = SendMessageLongA(m_lHdrHwnd, HDM_GETITEMCOUNT, 0&, 0&)
    End If
End Property

Public Property Get ColumnDragLine() As Boolean
'/* [get] column drag line
    ColumnDragLine = m_bColumnDragLine
End Property

Public Property Let ColumnDragLine(ByVal PropVal As Boolean)
'/* [let] column drag line
    m_bColumnDragLine = PropVal
    PropertyChanged "ColumnDragLine"
End Property

Private Property Get ColumnDragging() As Boolean
'/* [get] enable column drag and drop
    ColumnDragging = m_bColumnDragging
End Property

Private Property Let ColumnDragging(ByVal PropVal As Boolean)
'/* [let] enable column drag and drop
    m_bColumnDragging = PropVal
    m_cSkinHeader.DragState = PropVal
End Property

Public Property Get ColumnFilters() As Boolean
Attribute ColumnFilters.VB_MemberFlags = "40"
'/* [get] get column filters
    ColumnFilters = m_bColumnFilters
End Property

Public Property Let ColumnFilters(ByVal PropVal As Boolean)
'/* [let] enable column filters
    m_bColumnFilters = PropVal
    If Not (m_cSkinHeader Is Nothing) Then
        m_cSkinHeader.ColumnFilters = PropVal
        Resize
    End If
End Property

Private Property Get ColumnSizingVertical() As Boolean
'/* [get] enable vertical column sizing
    ColumnSizingVertical = m_bColumnSizingVertical
End Property

Private Property Let ColumnSizingVertical(ByVal PropVal As Boolean)
'/* [let] enable vertical column sizing
    m_bColumnSizingVertical = PropVal
End Property

Private Property Get ColumnSizingHorizontal() As Boolean
'/* [get]  enable horizontal column sizing
    ColumnSizingHorizontal = m_bColumnSizingHorizontal
End Property

Private Property Let ColumnSizingHorizontal(ByVal PropVal As Boolean)
'/* [let] enable horizontal column sizing
    m_bColumnSizingHorizontal = PropVal
End Property

Public Property Get ColumnFocus() As Boolean
'/* [get] enable column focus
    ColumnFocus = m_bColumnFocus
End Property

Public Property Let ColumnFocus(ByVal PropVal As Boolean)
'/* [let] enable column focus
    m_bColumnFocus = PropVal
    PropertyChanged "ColumnFocus"
End Property

Public Property Get ColumnFocusColor() As OLE_COLOR
'/* [get] column focus forecolor
    ColumnFocusColor = m_oColumnFocusColor
End Property

Public Property Let ColumnFocusColor(ByVal PropVal As OLE_COLOR)
'/* [let] column focus forecolor
    m_oColumnFocusColor = PropVal
    PropertyChanged "ColumnFocusColor"
End Property

Public Property Get ColumnHeight() As Long
Attribute ColumnHeight.VB_MemberFlags = "40"
'*/ [get] retrieve a columns height

Dim tHDR As RECT

    If Not (m_lHdrHwnd = 0) Then
        '/* get coordinates
        GetClientRect m_lHdrHwnd, tHDR
        ColumnHeight = tHDR.Bottom
    End If

End Property

Public Property Get ColumnIcon(ByVal lColumn As Long) As Long
Attribute ColumnIcon.VB_MemberFlags = "40"
'*/ [get] retieve header icon index

Dim uHDI As HDITEM

    If Not (m_lHdrHwnd = 0) Then
        If Not (m_lImlHdHndl = 0) Then
            With uHDI
                .Mask = HDI_FORMAT
                If m_bIsNt Then
                    SendMessageW m_lHdrHwnd, HDM_GETITEMW, lColumn, uHDI
                Else
                    SendMessageA m_lHdrHwnd, HDM_GETITEMA, lColumn, uHDI
                End If
                If (.fmt And HDF_IMAGE) = HDF_IMAGE Then
                    .Mask = HDI_IMAGE
                    If m_bIsNt Then
                        SendMessageW m_lHdrHwnd, HDM_GETITEMW, lColumn, uHDI
                    Else
                        SendMessageA m_lHdrHwnd, HDM_GETITEMA, lColumn, uHDI
                    End If
                    ColumnIcon = .iImage
                Else
                    ColumnIcon = -1
                End If
            End With
        End If
    End If

End Property

Public Property Let ColumnIcon(ByVal lColumn As Long, _
                               ByVal lIcon As Long)
'*/ [let] change header icon

Dim lAlign  As Long
Dim uHDI    As HDITEM

    If Not (m_lHdrHwnd = 0) Then
        If Not (m_lImlHdHndl = 0) Then
            If (lColumn < m_lColumnCount) Then
                With uHDI
                    .Mask = HDI_FORMAT
                    If m_bIsNt Then
                        SendMessageW m_lHdrHwnd, HDM_GETITEMW, lColumn, uHDI
                    Else
                        SendMessageA m_lHdrHwnd, HDM_GETITEMA, lColumn, uHDI
                    End If
                    lAlign = &H3 And .fmt
                    .iImage = lIcon
                    .fmt = HDF_STRING Or lAlign Or HDF_IMAGE * -(lIcon > -1 And m_lImlHdHndl <> 0) Or HDF_BITMAP_ON_RIGHT
                    .Mask = HDI_IMAGE * -(lIcon > -1) Or HDI_FORMAT
                End With
                If m_bIsNt Then
                    SendMessageW m_lHdrHwnd, HDM_SETITEMW, lColumn, uHDI
                Else
                    SendMessageA m_lHdrHwnd, HDM_SETITEMA, lColumn, uHDI
                End If
            End If
        End If
    End If

End Property

Private Property Get ColumnIndex(ByVal lColumn As Long) As Long
'/* [get] column index

Dim tLVI As LVCOLUMN
    
    If Not (m_lHGHwnd = 0) Then
        If (lColumn < m_lColumnCount) Then
            With tLVI
                .Mask = LVCF_ORDER
                If Not (SendMessageA(m_lHGHwnd, LVM_GETCOLUMNA, lColumn, tLVI) = 0) Then
                    ColumnIndex = .iOrder
                End If
            End With
        End If
    End If

End Property

Public Property Get ColumnLock(ByVal lColumn As Long) As Boolean
Attribute ColumnLock.VB_MemberFlags = "40"
'/* [get] column locked
    If ArrayCheck(m_bColumnLock) Then
        ColumnLock = m_bColumnLock(lColumn)
    End If
End Property

Public Property Let ColumnLock(ByVal lColumn As Long, _
                               ByVal bState As Boolean)
'/* [let] column locked
    If Not (m_lHGHwnd = 0) Then
        If ArrayCheck(m_bColumnLock) Then
            m_bColumnLock(lColumn) = bState
        End If
        If Not (m_cSkinHeader Is Nothing) Then
            If (lColumn < m_lColumnCount) Then
                If Not (m_cSkinHeader Is Nothing) Then
                    m_cSkinHeader.ColumnLocked(lColumn) = bState
                End If
            End If
        End If
    End If
End Property

Public Property Get ColumnTag(ByVal lColumn As Long) As ECSColumnSortTags
Attribute ColumnTag.VB_MemberFlags = "40"
'/* [get] column sort tag
    If (lColumn < m_lColumnCount) Then
        If Not (c_ColumnTags Is Nothing) Then
            On Error Resume Next
            ColumnTag = c_ColumnTags.Item(CStr(lColumn))
            On Error GoTo 0
        End If
        If (ColumnTag < 0) Or (ColumnTag > 4) Then
            ColumnTag = -1
        End If
    End If
End Property

Private Property Get ColumnText(ByVal lColumn As Long) As String
'*/ [get] column caption

Dim lLen       As Long
Dim aText(261) As Byte
Dim uLVC       As LVCOLUMN

    If Not (m_lHGHwnd = 0) Then
        If (lColumn < m_lColumnCount) Then
            If m_bIsNt Then
                With uLVC
                    .pszText = VarPtr(aText(0))
                    .cchTextMax = (UBound(aText) + 1)
                    .Mask = LVCF_TEXT
                    SendMessageW m_lHGHwnd, LVM_GETCOLUMNW, lColumn, uLVC
                    ColumnText = PointerToString(.pszText)
                End With
            Else
                With uLVC
                    .pszText = VarPtr(aText(0))
                    .cchTextMax = UBound(aText)
                    .Mask = LVCF_TEXT
                End With
                SendMessageA m_lHGHwnd, LVM_GETCOLUMNA, lColumn, uLVC
                ColumnText = StrConv(aText(), vbUnicode)
                lLen = InStr(ColumnText, vbNullChar)
                If (lLen > 0) Then
                    ColumnText = left$(ColumnText, lLen - 1)
                End If
            End If
        End If
    End If

End Property

Public Property Let ColumnText(ByVal lColumn As Long, _
                               ByVal sText As String)

'*/ [let] change a columns caption

Dim uLVC As LVCOLUMN

    If Not (m_lHGHwnd = 0) Then
        If (lColumn < m_lColumnCount) Then
            If m_bIsNt Then
                With uLVC
                    .pszText = StrPtr(sText)
                    .cchTextMax = LenB(sText)
                    .Mask = LVCF_TEXT
                End With
                SendMessageW m_lHGHwnd, LVM_SETCOLUMNW, lColumn, uLVC
            Else
                With uLVC
                    .pszText = sText
                    .cchTextMax = LenB(sText)
                    .Mask = LVCF_TEXT
                End With
                SendMessageA m_lHGHwnd, LVM_SETCOLUMNA, lColumn, uLVC
            End If
        End If
    End If

End Property

Public Property Get ColumnToolTips() As Boolean
Attribute ColumnToolTips.VB_MemberFlags = "40"
'*/ [get] enable tool tips
    If Not (m_cSkinHeader Is Nothing) Then
        ColumnToolTips = m_cSkinHeader.ToolTips
    End If
End Property

Public Property Let ColumnToolTips(ByVal PropVal As Boolean)
'*/ [let] enable tool tips
    If Not (m_cSkinHeader Is Nothing) Then
        m_cSkinHeader.ToolTips = PropVal
    End If
End Property

Public Property Get ColumnTipColor() As Long
Attribute ColumnTipColor.VB_MemberFlags = "40"
'*/ [get] tip backcolor
    If Not (m_cSkinHeader Is Nothing) Then
        ColumnTipColor = m_cSkinHeader.TipColor
    End If
End Property

Public Property Let ColumnTipColor(ByVal PropVal As Long)
'*/ [let] tip backcolor
    If Not (m_cSkinHeader Is Nothing) Then
        m_cSkinHeader.TipColor = PropVal
    End If
End Property

Public Property Get ColumnTipForeColor() As Long
Attribute ColumnTipForeColor.VB_MemberFlags = "40"
'*/ [get] tip gradient offset color
    If Not (m_cSkinHeader Is Nothing) Then
        ColumnTipForeColor = m_cSkinHeader.TipForeColor
    End If
End Property

Public Property Let ColumnTipForeColor(ByVal PropVal As Long)
'*/ [let] tip gradient offset color
    If Not (m_cSkinHeader Is Nothing) Then
        m_cSkinHeader.TipForeColor = PropVal
    End If
End Property

Public Property Get ColumnTipHint(ByVal lColumn As Long) As String
Attribute ColumnTipHint.VB_MemberFlags = "40"
'*/ [get] tip main caption
    ColumnTipHint = m_cSkinHeader.TipHint(lColumn)
End Property

Public Property Let ColumnTipHint(ByVal lColumn As Long, _
                                  ByVal sHint As String)
'*/ [let] tip main caption
    If Not (m_cSkinHeader Is Nothing) Then
        m_cSkinHeader.TipHint(lColumn) = sHint
    End If
End Property

Public Property Get ColumnTipDelayTime() As Long
Attribute ColumnTipDelayTime.VB_MemberFlags = "40"
'*/ [get] tip delay time
    If Not (m_cSkinHeader Is Nothing) Then
        ColumnTipDelayTime = m_cSkinHeader.TipDelayTime
    End If
End Property

Public Property Let ColumnTipDelayTime(ByVal PropVal As Long)
'*/ [let] tip delay time
    m_cSkinHeader.TipDelayTime = PropVal
End Property

Public Property Get ColumnTipFont() As StdFont
Attribute ColumnTipFont.VB_MemberFlags = "40"
'*/ [get] tip font
    If Not (m_cSkinHeader Is Nothing) Then
        Set ColumnTipFont = m_cSkinHeader.TipFont
    End If
End Property

Public Property Set ColumnTipFont(ByVal PropVal As StdFont)
'*/ [set] tip font
    If Not (m_cSkinHeader Is Nothing) Then
        Set m_cSkinHeader.TipFont = PropVal
    End If
End Property

Public Property Get ColumnTipGradient() As Boolean
Attribute ColumnTipGradient.VB_MemberFlags = "40"
'*/ [get] enable tip gradient
    If Not (m_cSkinHeader Is Nothing) Then
        ColumnTipGradient = m_cSkinHeader.TipGradient
    End If
End Property

Public Property Let ColumnTipGradient(ByVal PropVal As Boolean)
'*/ [let] enable tip gradient
    If Not (m_cSkinHeader Is Nothing) Then
        m_cSkinHeader.TipGradient = PropVal
    End If
End Property

Public Property Get ColumnTipMultiline() As Boolean
Attribute ColumnTipMultiline.VB_MemberFlags = "40"
'*/ [get] tip multiline
    If Not (m_cSkinHeader Is Nothing) Then
        ColumnTipMultiline = m_cSkinHeader.TipMultiline
    End If
End Property

Public Property Let ColumnTipMultiline(ByVal PropVal As Boolean)
'*/ [let] tip multiline
    If Not (m_cSkinHeader Is Nothing) Then
        m_cSkinHeader.TipMultiline = PropVal
    End If
End Property

Public Property Get ColumnTipOffsetColor() As Long
Attribute ColumnTipOffsetColor.VB_MemberFlags = "40"
'*/ [get] tip gradient offset color
    If Not (m_cSkinHeader Is Nothing) Then
        ColumnTipOffsetColor = m_cSkinHeader.TipOffsetColor
    End If
End Property

Public Property Let ColumnTipOffsetColor(ByVal PropVal As Long)
'*/ [let] tip gradient offset color
    If Not (m_cSkinHeader Is Nothing) Then
        m_cSkinHeader.TipOffsetColor = PropVal
    End If
End Property

Public Property Get ColumnTipPosition() As Long
Attribute ColumnTipPosition.VB_MemberFlags = "40"
'*/ [get] tip position
    If Not (m_cSkinHeader Is Nothing) Then
        ColumnTipPosition = m_cSkinHeader.TipPosition
    End If
End Property

Public Property Let ColumnTipPosition(ByVal PropVal As Long)
'*/ [let] tip position
    If Not (m_cSkinHeader Is Nothing) Then
        m_cSkinHeader.TipPosition = PropVal
    End If
End Property

Public Property Get ColumnTipTransparency() As Long
Attribute ColumnTipTransparency.VB_MemberFlags = "40"
'*/ [get] tip transparency
    If Not (m_cSkinHeader Is Nothing) Then
        ColumnTipTransparency = m_cSkinHeader.TipTransparency
    End If
End Property

Public Property Let ColumnTipTransparency(ByVal PropVal As Long)
'*/ [let] tip transparency
    If Not (m_cSkinHeader Is Nothing) Then
        m_cSkinHeader.TipTransparency = PropVal
    End If
End Property

Public Property Get ColumnTipVisibleTime() As Long
Attribute ColumnTipVisibleTime.VB_MemberFlags = "40"
'*/ [get] tip visible time
    If Not (m_cSkinHeader Is Nothing) Then
        ColumnTipVisibleTime = m_cSkinHeader.TipVisibleTime
    End If
End Property

Public Property Let ColumnTipVisibleTime(ByVal PropVal As Long)
'*/ [let] tip visible time
    If Not (m_cSkinHeader Is Nothing) Then
        m_cSkinHeader.TipVisibleTime = PropVal
    End If
End Property

Public Property Get ColumnTipXPColors() As Boolean
Attribute ColumnTipXPColors.VB_MemberFlags = "40"
'*/ [get] tip use xp color offsets
    If Not (m_cSkinHeader Is Nothing) Then
        ColumnTipXPColors = m_cSkinHeader.TipXPColors
    End If
End Property

Public Property Let ColumnTipXPColors(ByVal PropVal As Boolean)
'*/ [let] tip use xp color offsets
    If Not (m_cSkinHeader Is Nothing) Then
        m_cSkinHeader.TipXPColors = PropVal
    End If
End Property

Public Property Get ColumnVerticalText() As Boolean
Attribute ColumnVerticalText.VB_MemberFlags = "40"
'*/ [get] enable column vertical text
    If m_bSkinHeader Then
        ColumnVerticalText = m_cSkinHeader.ColumnVerticalText
    End If
    ColumnVerticalText = m_bColumnVerticalText
End Property

Public Property Let ColumnVerticalText(ByVal PropVal As Boolean)
'*/ [let] enable column vertical text
    If m_bSkinHeader Then
        m_cSkinHeader.ColumnVerticalText = PropVal
    End If
    m_bColumnVerticalText = PropVal
End Property

Public Property Get ColumnWidth(ByVal lColumn As Long) As Long
Attribute ColumnWidth.VB_MemberFlags = "40"
'*/ [get] retrieve a columns length
    If Not (m_lHGHwnd = 0) Then
        If (lColumn < m_lColumnCount) Then
            ColumnWidth = SendMessageLongA(m_lHGHwnd, LVM_GETCOLUMNWIDTH, lColumn, 0&)
        End If
    End If
End Property

Public Property Let ColumnWidth(ByVal lColumn As Long, _
                                ByVal lWidth As Long)

'*/ [let] change a columns length
    If Not (m_lHGHwnd = 0) Then
        If (lColumn < m_lColumnCount) Then
            SendMessageLongA m_lHGHwnd, LVM_SETCOLUMNWIDTH, lColumn, lWidth
        End If
    End If
End Property

Public Property Get Count() As Long
Attribute Count.VB_MemberFlags = "40"
'*/ [get] row count
    If Not (m_lHGHwnd = 0) Then
        Count = SendMessageLongA(m_lHGHwnd, LVM_GETITEMCOUNT, 0&, 0&)
    End If
End Property

Public Property Get CustomCursors() As Boolean
'*/ [get] custom cursors
    CustomCursors = m_bCustomCursors
End Property

Public Property Let CustomCursors(ByVal PropVal As Boolean)
'*/ [let] custom cursors
    If Not (m_cSkinHeader Is Nothing) Then
        m_cSkinHeader.CustomCursors = PropVal
    End If
    m_bCustomCursors = PropVal
    PropertyChanged "CustomCursors"
End Property

Public Property Get DisabledBackColor() As Long
Attribute DisabledBackColor.VB_MemberFlags = "40"
'*/ [get] disabled backcolor
    DisabledBackColor = m_lDisabledBackColor
End Property

Public Property Let DisabledBackColor(ByVal PropVal As Long)
'*/ [let] disabled backcolor
    m_lDisabledBackColor = PropVal
End Property

Public Property Get DisabledForeColor() As Long
Attribute DisabledForeColor.VB_MemberFlags = "40"
'*/ [get] disabled forecolor
    DisabledForeColor = m_lDisabledForeColor
End Property

Public Property Let DisabledForeColor(ByVal PropVal As Long)
'*/ [let] disabled forecolor
    m_lDisabledForeColor = PropVal
End Property

Public Property Get DragEffectStyle() As EDSDragEffectStyle
'/* [get] drag decoration style
    DragEffectStyle = m_eDragEffectStyle
End Property

Public Property Let DragEffectStyle(ByVal PropVal As EDSDragEffectStyle)
'/* [let] drag decoration style
    m_eDragEffectStyle = PropVal
    PropertyChanged "DragEffectStyle"
End Property

Public Property Get DoubleBuffer() As Boolean
'*/ [get] buffer draw phase
    DoubleBuffer = m_bDoubleBuffer
End Property

Public Property Let DoubleBuffer(ByVal PropVal As Boolean)
'*/ [let] buffer draw phase
    If PropVal Then
        If Not (m_cGridBuffer Is Nothing) Then
            Set m_cGridBuffer = Nothing
        End If
        Set m_cGridBuffer = New clsStoreDc
    Else
        Set m_cGridBuffer = Nothing
    End If
    m_bDoubleBuffer = PropVal
    PropertyChanged "DoubleBuffer"
End Property

Public Property Get Draw() As Boolean
Attribute Draw.VB_MemberFlags = "40"
'*/ [get] enable drawing
    Draw = m_bDraw
End Property

Public Property Let Draw(ByVal PropVal As Boolean)
'*/ [let] enable drawing
    m_bDraw = PropVal
End Property

Public Property Get EditBlendBackground() As Boolean
Attribute EditBlendBackground.VB_MemberFlags = "40"
'*/ [get] blend edit box back ground
    EditBlendBackground = m_bEditBlendBackground
End Property

Public Property Let EditBlendBackground(ByVal PropVal As Boolean)
'*/ [let] blend edit box back ground
    m_bEditBlendBackground = PropVal
End Property

Public Property Get EditFrameStyle() As EBSBorderStyle
Attribute EditFrameStyle.VB_MemberFlags = "40"
'/* [get] edit box frame style
    EditFrameStyle = m_eEditFrameStyle
End Property

Public Property Let EditFrameStyle(ByVal PropVal As EBSBorderStyle)
'*/ [let] edit box frame style
    m_eEditFrameStyle = PropVal
End Property

Public Property Get Enabled() As Boolean
'/* [get] toggle listview enable state
    Enabled = m_bEnabled
End Property

Public Property Let Enabled(ByVal PropVal As Boolean)
'/* [let] toggle listview enable state
    If Not (m_lHGHwnd = 0) Then
        If PropVal Then
            If Not m_bEnabled Then
                SendMessageLongA m_lHGHwnd, LVM_SETBKCOLOR, 0&, m_oBackColor
                SendMessageLongA m_lHGHwnd, LVM_SETTEXTBKCOLOR, 0&, m_oBackColor
                SendMessageLongA m_lHGHwnd, LVM_SETTEXTCOLOR, 0&, m_oForeColor
                EnableWindow m_lHGHwnd, 1&
            End If
        Else
            If m_bEnabled Then
                SendMessageLongA m_lHGHwnd, LVM_SETBKCOLOR, 0&, m_lDisabledBackColor
                SendMessageLongA m_lHGHwnd, LVM_SETTEXTBKCOLOR, 0&, m_lDisabledBackColor
                SendMessageLongA m_lHGHwnd, LVM_SETTEXTCOLOR, 0&, m_lDisabledForeColor
                EnableWindow m_lHGHwnd, 0&
            End If
        End If
        RaiseEvent eVHGridEnable(PropVal)
    End If
    If Not (m_cSkinScrollBars Is Nothing) Then
        m_cSkinScrollBars.Enabled = PropVal
    End If
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.Enabled = PropVal
    End If
    m_bEnabled = PropVal
    PropertyChanged "Enabled"
End Property

Public Property Get FastLoad() As Boolean
Attribute FastLoad.VB_MemberFlags = "40"
'*/ [get] post init drawing
    FastLoad = m_bFastLoad
End Property

Public Property Let FastLoad(ByVal PropVal As Boolean)
'*/ [let] post init drawing
    If PropVal Then
        m_bDraw = False
    End If
    m_bFastLoad = PropVal
End Property

Public Property Get FilterBackColor() As OLE_COLOR
Attribute FilterBackColor.VB_MemberFlags = "40"
'*/ [get] filter box back color
    FilterBackColor = m_oFilterBackColor
End Property

Public Property Let FilterBackColor(ByVal PropVal As OLE_COLOR)
'*/ [let] filter box back color
    m_oFilterBackColor = PropVal
End Property

Public Property Get FilterControlColor() As OLE_COLOR
Attribute FilterControlColor.VB_MemberFlags = "40"
'*/ [get] filter box controls color
    FilterControlColor = m_oFilterControlColor
End Property

Public Property Let FilterControlColor(ByVal PropVal As OLE_COLOR)
'*/ [let] filter box controls color
    m_oFilterControlColor = PropVal
End Property

Public Property Get FilterControlForeColor() As OLE_COLOR
Attribute FilterControlForeColor.VB_MemberFlags = "40"
'*/ [get] filter box controls color
    FilterControlForeColor = m_oFilterControlForeColor
End Property

Public Property Let FilterControlForeColor(ByVal PropVal As OLE_COLOR)
'*/ [let] filter box controls color
    m_oFilterControlForeColor = PropVal
End Property

Public Property Get FilterForeColor() As OLE_COLOR
Attribute FilterForeColor.VB_MemberFlags = "40"
'*/ [get] filter box forecolor
    FilterForeColor = m_oFilterForeColor
End Property

Public Property Let FilterForeColor(ByVal PropVal As OLE_COLOR)
'*/ [let] filter box forecolor
    m_oFilterForeColor = PropVal
End Property

Public Property Get FilterTitleColor() As OLE_COLOR
Attribute FilterTitleColor.VB_MemberFlags = "40"
'*/ [get] filter box title color
    FilterTitleColor = m_oFilterTitleColor
End Property

Public Property Let FilterTitleColor(ByVal PropVal As OLE_COLOR)
'*/ [let] filter box title color
    m_oFilterTitleColor = PropVal
End Property

Public Property Get FilterGradient() As Boolean
Attribute FilterGradient.VB_MemberFlags = "40"
'*/ [get] enable filter gradient
    FilterGradient = m_bFilterGradient
End Property

Public Property Let FilterGradient(ByVal PropVal As Boolean)
'*/ [let] enable filter gradient
    m_bFilterGradient = PropVal
End Property

Private Property Get FilterLoaded() As Boolean
'*/ [get] filter loaded state
    FilterLoaded = m_bFilterLoaded
End Property

Private Property Let FilterLoaded(ByVal PropVal As Boolean)
'*/ [let] filter loaded state
    m_cSkinHeader.FilterLoaded = PropVal
    m_bFilterLoaded = PropVal
End Property

Public Property Get FilterOffsetColor() As OLE_COLOR
Attribute FilterOffsetColor.VB_MemberFlags = "40"
'*/ [get] filter gradient offset color
    FilterOffsetColor = m_oFilterOffsetColor
End Property

Public Property Let FilterOffsetColor(ByVal PropVal As OLE_COLOR)
'*/ [let] filter gradient offset color
    m_oFilterOffsetColor = PropVal
End Property

Public Property Get FilterTransparency() As Long
Attribute FilterTransparency.VB_MemberFlags = "40"
'*/ [get] enable filter box transparency
    FilterTransparency = m_lFilterTransparency
End Property

Public Property Let FilterTransparency(ByVal PropVal As Long)
'*/ [let] enable filter box transparency
    m_lFilterTransparency = PropVal
End Property

Public Property Get FilterXPColors() As Boolean
Attribute FilterXPColors.VB_MemberFlags = "40"
'*/ [get] filter box use xp color offsets
    FilterXPColors = m_bFilterXPColors
End Property

Public Property Let FilterXPColors(ByVal PropVal As Boolean)
'*/ [let] filter box use xp color offsets
    m_bFilterXPColors = PropVal
End Property

Public Property Get GridFocus() As Boolean
Attribute GridFocus.VB_MemberFlags = "40"
'/* [get] grid focus
    GridFocus = (GetFocus() = m_lHGHwnd)
End Property

Public Property Let GridFocus(ByVal PropVal As Boolean)
'/* [let] grid focus
    If Not (m_lHGHwnd = 0) Then
        If m_bIsNt Then
            If PropVal Then
                PostMessageW m_lHGHwnd, WM_SETFOCUS, 0&, 0&
            Else
                PostMessageW m_lHGHwnd, WM_KILLFOCUS, 0&, 0&
            End If
        Else
            If PropVal Then
                PostMessageA m_lHGHwnd, WM_SETFOCUS, 0&, 0&
            Else
                PostMessageA m_lHGHwnd, WM_KILLFOCUS, 0&, 0&
            End If
        End If
    End If
End Property

Public Property Get FirstRowReserved() As Boolean
Attribute FirstRowReserved.VB_MemberFlags = "40"
'/* [get] no focus on first row
    FirstRowReserved = m_bFirstRowReserved
End Property

Public Property Let FirstRowReserved(ByVal PropVal As Boolean)
'/* [let] no focus on first row
    m_bFirstRowReserved = PropVal
End Property

Public Property Get FocusTextOnly() As Boolean
'/* [get] focus box on text only
    FocusTextOnly = m_bFocusTextOnly
End Property

Public Property Let FocusTextOnly(ByVal PropVal As Boolean)
'/* [let] focus box on text only
    m_bFocusTextOnly = PropVal
    PropertyChanged "FocusTextOnly"
End Property

Public Property Get FocusAlphaBlend() As Boolean
'/* [get] alpha blend focus color
    FocusAlphaBlend = m_bAlphaBlend
End Property

Public Property Let FocusAlphaBlend(ByVal PropvVal As Boolean)
'/* [let] alpha blend focus color
    m_bAlphaBlend = PropvVal
    PropertyChanged "FocusAlphaBlend"
End Property

Public Property Get FontRightLeading() As Boolean
Attribute FontRightLeading.VB_MemberFlags = "40"
'/* [get] right align fonts
    FontRightLeading = m_bFontRightLeading
End Property

Public Property Let FontRightLeading(ByVal PropvVal As Boolean)
'/* [let] right align fonts
    m_bFontRightLeading = PropvVal
End Property

Public Property Get Font() As StdFont
'/* [get] retrieve list font
    If Not (m_oFont Is Nothing) Then
        Set Font = m_oFont
    End If
End Property

Public Property Set Font(ByVal oFont As StdFont)
'*/ [set] change list font

Dim lChar       As Long
Dim uLF         As LOGFONT
Dim bteFont()   As Byte

    Set m_oFont = oFont
    If Not (m_lHGHwnd = 0) Then
        If Not (oFont Is Nothing) Then
            DestroyFont
            '/* extract properties
            With uLF
                bteFont = StrConv(oFont.Name, vbFromUnicode)
                For lChar = 0 To UBound(bteFont)
                    .lfFaceName(lChar) = bteFont(lChar)
                Next lChar
                .lfHeight = -MulDiv(oFont.Size, GetDeviceCaps(UserControl.hdc, LOGPIXELSY), 72)
                .lfItalic = oFont.Italic
                .lfWeight = IIf(oFont.Bold, FW_BOLD, FW_NORMAL)
                .lfUnderline = oFont.Underline
                .lfStrikeOut = oFont.Strikethrough
                If m_bUseUnicode Then
                    .lfCharSet = 134
                Else
                    .lfCharSet = 3
                End If
                If m_bIsXp Then
                    .lfQuality = LF_CLEARTYPE_QUALITY
                Else
                    .lfQuality = LF_ANTIALIASED_QUALITY
                End If
            End With
            '/* create the log font
            If m_bUseUnicode Then
                m_lFont = CreateFontIndirectW(uLF)
                SendMessageLongW m_lHGHwnd, WM_SETFONT, m_lFont, True
            Else
                m_lFont = CreateFontIndirectA(uLF)
                SendMessageLongA m_lHGHwnd, WM_SETFONT, m_lFont, True
            End If
        End If
    End If
    PropertyChanged "Font"

End Property

Public Property Get ForeColor() As OLE_COLOR
'*/ [get] retrieve list forecolor
    ForeColor = m_oForeColor
End Property

Public Property Let ForeColor(ByVal PropVal As OLE_COLOR)
'*/ [let] change list forecolor
    OleTranslateColor PropVal, 0&, m_oForeColor
    m_oForeColor = PropVal
    If Not (m_lHGHwnd = 0) Then
        If m_bEnabled Then
            SendMessageLongA m_lHGHwnd, LVM_SETTEXTCOLOR, 0&, m_oBackColor
        End If
    End If
    PropertyChanged "ForeColor"
End Property

Public Property Get ForeColorFocused() As OLE_COLOR
'*/ [get] retrieve list focused forecolor
    ForeColorFocused = m_oCellFocusedHighlight
End Property

Public Property Let ForeColorFocused(ByVal PropVal As OLE_COLOR)
'*/ [let] change list focused forecolor
    If Not (PropVal < 0) Then
        OleTranslateColor PropVal, 0&, m_oCellFocusedHighlight
        m_oCellFocusedHighlight = PropVal
        PropertyChanged "ForeColorFocused"
    End If
End Property

Public Property Get ForeColorAuto() As Boolean
Attribute ForeColorAuto.VB_MemberFlags = "40"
'*/ [get] auto focus forecolor
    ForeColorAuto = m_bForeColorAuto
End Property

Public Property Let ForeColorAuto(ByVal PropVal As Boolean)
'*/ [let] auto focus forecolor
    m_bForeColorAuto = PropVal
End Property

Public Property Get FullRowSelect() As Boolean
'*/ [get] retrieve full row select state
    FullRowSelect = m_bFullRowSelect
End Property

Public Property Let FullRowSelect(ByVal PropVal As Boolean)
'*/ [let] change full row select state
    m_bFullRowSelect = PropVal
    PropertyChanged "FullRowSelect"
End Property

Public Property Get GridLines() As EGLGridLines
'*/ [get] gridlines style
    GridLines = m_eGridLines
End Property

Public Property Let GridLines(ByVal PropVal As EGLGridLines)
'*/ [let] gridlines style
    m_eGridLines = PropVal
    PropertyChanged "GridLines"
End Property

Public Property Get GridLineColor() As OLE_COLOR
'*/ [get] gridline color
    GridLineColor = m_oGridLineColor
End Property

Public Property Let GridLineColor(ByVal PropVal As OLE_COLOR)
'*/ [let] gridline color
    m_oGridLineColor = PropVal
    PropertyChanged "GridLineColor"
End Property

Public Property Get GridInitialized() As Boolean
Attribute GridInitialized.VB_MemberFlags = "40"
'*/ [get] grid loaded state
    GridInitialized = m_bHasInitialized
End Property

Public Property Let GridInitialized(ByVal PropVal As Boolean)
'*/ [let] grid loaded state
    m_bHasInitialized = PropVal
End Property

Public Property Get HeaderDragDrop() As Boolean
'*/ [get] retrieve drag and drop state
    HeaderDragDrop = m_bDragDrop
End Property

Public Property Let HeaderDragDrop(ByVal PropVal As Boolean)
'*/ [let] retrieve drag and drop state
    If Not (m_lHGHwnd = 0) Then
        If PropVal Then
            SetExtendedStyle LVS_EX_HEADERDRAGDROP, 0
        Else
            SetExtendedStyle 0, LVS_EX_HEADERDRAGDROP
        End If
    End If
    m_bDragDrop = PropVal
    PropertyChanged "HeaderDragDrop"
End Property

Public Property Get HeaderFixedWidth() As Boolean
'*/ [get] retrieve fixed width state
    HeaderFixedWidth = m_bHeaderFixed
End Property

Public Property Let HeaderFixedWidth(ByVal PropVal As Boolean)
'*/ [let] change fixed width state
    If Not (m_cSkinHeader Is Nothing) Then
        m_cSkinHeader.HeaderFixedWidth = PropVal
    End If
    m_bHeaderFixed = PropVal
    PropertyChanged "HeaderFixedWidth"
End Property

Public Property Get HeaderFlat() As Boolean
Attribute HeaderFlat.VB_MemberFlags = "40"
'*/ [get] retrieve flat header state
    HeaderFlat = m_bHeaderFlat
End Property

Public Property Let HeaderFlat(ByVal PropVal As Boolean)
'*/ [let] change flat header state
    If Not (m_cSkinHeader Is Nothing) Then
        m_cSkinHeader.HeaderFlat = PropVal
    End If
    m_bHeaderFlat = PropVal
End Property

Public Property Get HeaderFont() As StdFont
'/* [get] header font
    Set HeaderFont = m_oHeaderFont
End Property

Public Property Set HeaderFont(ByVal PropVal As StdFont)
'/* [set] header font
    Set m_oHeaderFont = PropVal
    PropertyChanged "HeaderFont"
End Property

Public Property Get HeaderForeColor() As OLE_COLOR
'/* [get] return the header forecolor
    HeaderForeColor = m_oHdrForeClr
End Property

Public Property Let HeaderForeColor(ByVal PropVal As OLE_COLOR)
'/* [let] change the header forecolor
    If Not (m_cSkinHeader Is Nothing) Then
        m_cSkinHeader.HeaderForeColor = PropVal
    End If
    m_oHdrForeClr = PropVal
    PropertyChanged "HeaderForeColor"
End Property

Public Property Get HeaderHeight() As Long
'/* [get] header height
    HeaderHeight = m_lHeaderHeight
End Property

Public Property Let HeaderHeight(ByVal PropVal As Long)
'/* [let] header height

Dim tRect As RECT

    If (PropVal < HDR_MINHEIGHT) Then
        If (m_lHeaderHeight > PropVal) Then
            Exit Property
        End If
    ElseIf (PropVal > HDR_MAXHEIGHT) Then
        If (m_lHeaderHeight < PropVal) Then
            Exit Property
        End If
    End If
    '/* send size change call, will change in proc
    If Not (m_lHdrHwnd = 0) Then
        m_lHeaderHeight = PropVal
        GetClientRect m_lHdrHwnd, tRect
        With tRect
            SetWindowPos m_lHdrHwnd, 0&, -2&, 0&, .Right, m_lHeaderHeight, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOZORDER
        End With
        '/* force a  complete redraw
        EraseRect m_lHdrHwnd, tRect, 0&
        If m_bHasInitialized Then
            If Not (m_lRowCount = 0) Then
                Dim lHdc As Long
                lHdc = GetDC(m_lHGHwnd)
                DrawTransitionMask lHdc
                ReleaseDC m_lHGHwnd, lHdc
            End If
        End If
        If Not m_bTransitionMask Then
            SetRowCount Count
        End If
    End If
    If (m_lHeaderHeight > 0) Then
        m_lHeaderOffset = m_lHeaderHeight + 4
    Else
        m_lHeaderOffset = 2
    End If
    PropertyChanged "HeaderHeight"

End Property

Public Property Get HeaderHeightSizable() As Boolean
'/* [get] headerheight sizable enable
    HeaderHeightSizable = m_bHeaderSizable
End Property

Public Property Let HeaderHeightSizable(ByVal PropVal As Boolean)
'/* [let] headerheight sizable enable
    If Not (m_cSkinHeader Is Nothing) Then
        m_cSkinHeader.HeaderSizeable = PropVal
    End If
    m_bHeaderSizable = PropVal
    PropertyChanged "HeaderHeightSizable"
End Property

Public Property Get HeaderHide() As Boolean
Attribute HeaderHide.VB_MemberFlags = "40"
'/* [get] retrieve header visible state
    HeaderHide = m_bHeaderHide
End Property

Public Property Let HeaderHide(ByVal PropVal As Boolean)
'/* [let] change header visible state

    If Not (m_lHGHwnd = 0) Then
        m_bHeaderHide = PropVal
        If PropVal Then
            m_lStoreHeaderHeight = m_lHeaderHeight
            m_lHeaderHeight = 0
            SetStyle LVS_NOCOLUMNHEADER, 0
        Else
            m_lHeaderHeight = m_lStoreHeaderHeight
            SetStyle 0, LVS_NOCOLUMNHEADER
        End If
    End If
    GridRefresh True

End Property

Public Property Get HeaderForeColorFocused() As OLE_COLOR
'*/ [get] return header focus color
    HeaderForeColorFocused = m_oHdrHighLiteClr
End Property

Public Property Let HeaderForeColorFocused(ByVal PropVal As OLE_COLOR)
'*/ [let] change header focus color
    If Not (m_cSkinHeader Is Nothing) Then
        m_cSkinHeader.HeaderHighLite = PropVal
    End If
    m_oHdrHighLiteClr = PropVal
    PropertyChanged "HeaderForeColorFocused"
End Property

Private Function HeaderHwnd() As Long
'*/ return the column header handle
    If Not (m_lHGHwnd = 0) Then
        m_lHdrHwnd = SendMessageLongA(m_lHGHwnd, LVM_GETHEADER, 0&, 0&)
        HeaderHwnd = m_lHdrHwnd
    End If
End Function

Public Property Get HeaderForeColorPressed() As OLE_COLOR
'*/ [get] return header highlite color
    HeaderForeColorPressed = m_oHdrPressedClr
End Property

Public Property Let HeaderForeColorPressed(ByVal PropVal As OLE_COLOR)
'*/ [let] change header highlite color
    If Not (m_cSkinHeader Is Nothing) Then
        m_cSkinHeader.HeaderPressed = PropVal
    End If
    m_oHdrPressedClr = PropVal
    PropertyChanged "HeaderForeColorPressed"
End Property

Private Property Get HeaderHitState() As EHdrHitTest
    If Not (m_cSkinHeader Is Nothing) Then
        HeaderHitState = m_cSkinHeader.HeaderLastHitState
    End If
End Property

Private Property Let HeaderHitState(ByVal PropVal As EHdrHitTest)
    If Not (m_cSkinHeader Is Nothing) Then
        m_cSkinHeader.HeaderLastHitState = PropVal
    End If
End Property

Private Property Get HeaderHitTest() As EHdrHitTest
    If Not (m_cSkinHeader Is Nothing) Then
        HeaderHitTest = m_cSkinHeader.HeaderHitTest
    End If
End Property

Public Property Get HeaderTextEffect() As EHTTextEffect
Attribute HeaderTextEffect.VB_MemberFlags = "40"
'/* [get] header font effect
    HeaderTextEffect = m_lHeaderTextEffect
End Property

Public Property Let HeaderTextEffect(ByVal PropVal As EHTTextEffect)
'/* [let] header font effect
    m_lHeaderTextEffect = PropVal
End Property

Private Property Get HotTrackColor() As OLE_COLOR
'/* [get] cell hot track color
    HotTrackColor = m_oHotTrackColor
End Property

Private Property Let HotTrackColor(ByVal PropVal As OLE_COLOR)
'/* [let] cell hot track color
    m_oHotTrackColor = PropVal
End Property

Private Property Get HotTrackDepth() As ECHTrackDepth
'/* [get] hot track style depth
    HotTrackDepth = m_eHotTrackDepth
End Property

Private Property Let HotTrackDepth(ByVal PropVal As ECHTrackDepth)
'/* [let] hot track style depth
    m_eHotTrackDepth = PropVal
End Property

Public Property Get IconPosition() As Long
Attribute IconPosition.VB_MemberFlags = "40"
'/* [get] icon horizontal position
    IconPosition = m_lIconPosition
End Property

Public Property Let IconPosition(ByVal PropVal As ECPIconPosition)
'/* [let] icon horizontal position
    m_lIconPosition = PropVal
End Property

Public Property Get IconNoHilite() As Boolean
Attribute IconNoHilite.VB_MemberFlags = "40"
'/* [get] icon selected hilite
    IconNoHilite = m_bIconNoHilite
End Property

Public Property Let IconNoHilite(ByVal PropVal As Boolean)
'/* [let] icon selected hilite
    m_bIconNoHilite = PropVal
End Property

Private Property Get ISelectorBar() As StdPicture
'/* selector bar image
    Set ISelectorBar = m_pISelectorBar
End Property

Private Property Set ISelectorBar(ByVal PropVal As StdPicture)
'/* [set] alpha bar image
    Set m_pISelectorBar = PropVal
End Property

Public Property Get IsWinNT() As Boolean
Attribute IsWinNT.VB_MemberFlags = "40"
'/* [get] windows nt version
    IsWinNT = m_bIsNt
End Property

Public Property Get IsWinXP() As Boolean
Attribute IsWinXP.VB_MemberFlags = "40"
'/* [get] windows xp version
    IsWinXP = m_bIsXp
End Property

Public Property Get CellIndent(ByVal lRow As Long, _
                               ByVal lColumn As Long) As Long
Attribute CellIndent.VB_MemberFlags = "40"
'*/ [get] return row indent
    If Not (m_lHGHwnd = 0) Then
        If Not m_bVirtualMode Then
            If (lRow < m_lRowCount) Then
                If (lColumn < m_lColumnCount) Then
                    CellIndent = m_cGridItem(lRow).Indent(lColumn)
                End If
            End If
        End If
    End If
End Property

Public Property Let CellIndent(ByVal lRow As Long, _
                               ByVal lColumn As Long, _
                               ByVal lIndent As Long)

'*/ [let] change row indent
    If Not (m_lHGHwnd = 0) Then
        If Not m_bVirtualMode Then
            If (lRow < m_lRowCount) Then
                If (lColumn < m_lColumnCount) Then
                    m_cGridItem(lRow).Indent(lColumn) = lIndent
                End If
            End If
        End If
    End If
End Property

Public Property Get CellInFocus() As Long
Attribute CellInFocus.VB_MemberFlags = "40"
'*/ [get] focused cell
    CellInFocus = m_lCellFocused
End Property

Public Property Let CellInFocus(ByVal lCell As Long)
'*/ [let] focused cell
    m_lCellFocused = lCell
End Property

Public Property Get CellsSorted() As Boolean
Attribute CellsSorted.VB_MemberFlags = "40"
'*/ [get] return sorted mode status
    CellsSorted = m_bUseSorted
End Property

Public Property Let CellsSorted(ByVal PropVal As Boolean)
'*/ [let] change sorted mode status
    m_bUseSorted = PropVal
End Property

Public Property Get CellText(ByVal lRow As Long, _
                             ByVal lColumn As Long) As String
Attribute CellText.VB_MemberFlags = "40"

'*/ [get] return cell text
    If Not (m_lHGHwnd = 0) Then
        If Not m_bVirtualMode Then
            If (lRow < m_lRowCount) Then
                If (lColumn < m_lColumnCount) Then
                    CellText = m_cGridItem(lRow).Text(lColumn)
                End If
            End If
        End If
    End If
End Property

Public Property Let CellText(ByVal lRow As Long, _
                             ByVal lColumn As Long, _
                             ByVal sText As String)

'*/ [let] change row text
    If Not (m_lHGHwnd = 0) Then
        If Not m_bVirtualMode Then
            If (lRow < m_lRowCount) Then
                If (lColumn < m_lColumnCount) Then
                    m_cGridItem(lRow).Text(lColumn) = sText
                End If
            End If
        End If
    End If
End Property

Public Property Get CellTips() As Boolean
Attribute CellTips.VB_MemberFlags = "40"
'/* [get] enable cell tips
    CellTips = m_bCellTips
End Property

Public Property Let CellTips(ByVal PropVal As Boolean)
'/* [let] enable cell tips
    If Not m_bVirtualMode Then
        If PropVal Then
            If m_cCellTips Is Nothing Then
                Set m_cCellTips = New clsToolTip
            End If
            CellTipStart
        Else
            Set m_cCellTips = Nothing
        End If
        m_bCellTips = PropVal
    End If
End Property

Public Property Get CellTipColor() As OLE_COLOR
Attribute CellTipColor.VB_MemberFlags = "40"
'/* [get] cell tip color
    CellTipColor = m_oCellTipColor
End Property

Public Property Let CellTipColor(ByVal PropVal As OLE_COLOR)
'/* [let] cell tip color
    m_oCellTipColor = PropVal
End Property

Public Property Get CellTipDelayTime() As Long
Attribute CellTipDelayTime.VB_MemberFlags = "40"
'/* [get] cell tip delay time
    CellTipDelayTime = m_lCellTipDelayTime
End Property

Public Property Let CellTipDelayTime(ByVal PropVal As Long)
'/* [let]  cell tip delay time
    m_lCellTipDelayTime = PropVal
End Property

Public Property Get CellTipFont() As StdFont
Attribute CellTipFont.VB_MemberFlags = "40"
'/* [get] cell tip font
    Set CellTipFont = m_oCellTipFont
End Property

Public Property Set CellTipFont(ByVal PropVal As StdFont)
'/* [set] cell tip font
    Set m_oCellTipFont = PropVal
End Property

Public Property Get CellTipForeColor() As OLE_COLOR
Attribute CellTipForeColor.VB_MemberFlags = "40"
'/* [get] cell tip forecolor
    CellTipForeColor = m_oCellTipForeColor
End Property

Public Property Let CellTipForeColor(ByVal PropVal As OLE_COLOR)
'/* [let] cell tip forecolor
    m_oCellTipForeColor = PropVal
End Property

Public Property Get CellTipGradient() As Boolean
Attribute CellTipGradient.VB_MemberFlags = "40"
'/* [get] enable cell tip gradient
    CellTipGradient = m_bCellTipGradient
End Property

Public Property Let CellTipGradient(ByVal PropVal As Boolean)
'/* [let] enable cell tip gradient
    m_bCellTipGradient = PropVal
End Property

Public Property Get CellTipMultiline() As Boolean
Attribute CellTipMultiline.VB_MemberFlags = "40"
'/* [get] cell tip multiline
    CellTipMultiline = m_bCellTipMultiline
End Property

Public Property Let CellTipMultiline(ByVal PropVal As Boolean)
'/* [let] cell tip multiline
    m_bCellTipMultiline = PropVal
End Property

Public Property Get CellTipOffsetColor() As OLE_COLOR
Attribute CellTipOffsetColor.VB_MemberFlags = "40"
'/* [get] cell tip gradient offset
    CellTipOffsetColor = m_oCellTipOffsetColor
End Property

Public Property Let CellTipOffsetColor(ByVal PropVal As OLE_COLOR)
'/* [let] cell tip gradient offset
    m_oCellTipOffsetColor = PropVal
End Property

Public Property Get CellTipPosition() As ETTToolTipPosition
Attribute CellTipPosition.VB_MemberFlags = "40"
'/* [get] cell tip position
    CellTipPosition = m_lCellTipPosition
End Property

Public Property Let CellTipPosition(ByVal PropVal As ETTToolTipPosition)
'/* [let] cell tip position
    m_lCellTipPosition = PropVal
End Property

Public Property Get CellTipTransparency() As Long
Attribute CellTipTransparency.VB_MemberFlags = "40"
'/* [get] cell tip transparency
    CellTipTransparency = m_lCellTipTransparency
End Property

Public Property Let CellTipTransparency(ByVal PropVal As Long)
'/* [let] cell tip transparency
    m_lCellTipTransparency = PropVal
End Property

Public Property Get CellTipVisibleTime() As Long
Attribute CellTipVisibleTime.VB_MemberFlags = "40"
'/* [get] cell tip visible time
    CellTipVisibleTime = m_lCellTipVisibleTime
End Property

Public Property Let CellTipVisibleTime(ByVal PropVal As Long)
'/* [let] cell tip visible time
    m_lCellTipVisibleTime = PropVal
End Property

Public Property Get CellTipXPColors() As Boolean
Attribute CellTipXPColors.VB_MemberFlags = "40"
'/* [get] cell tip xp color offsets
    CellTipXPColors = m_bCellTipXPColors
End Property

Public Property Let CellTipXPColors(ByVal PropVal As Boolean)
'/* [let] cell tip xp color offsets
    m_bCellTipXPColors = PropVal
End Property

Public Property Get LockFirstColumn() As Boolean
Attribute LockFirstColumn.VB_MemberFlags = "40"
'*/ [get] lock the first column width
    LockFirstColumn = m_bLockFirstColumn
End Property

Public Property Let LockFirstColumn(ByVal PropVal As Boolean)
'*/ [let] lock the first column width
    If (PropVal < m_lColumnCount) Then
        If ArrayCheck(m_bColumnLock) Then
            ColumnLock(0) = PropVal
            m_bLockFirstColumn = PropVal
        End If
    End If
End Property

Public Property Get OLEDragMode() As OLEDragConstants
'*/ [get] ole drag mode
    OLEDragMode = m_eOLEDragMode
End Property

Public Property Let OLEDragMode(ByVal PropVal As OLEDragConstants)
'*/ [let] ole drag mode
    m_eOLEDragMode = PropVal
    PropertyChanged "OLEDragMode"
End Property

Public Property Get OLEDropMode() As EDCDropConstants
'*/ [get] ole drop mode
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal PropVal As EDCDropConstants)
'*/ [let] ole drop mode
    UserControl.OLEDropMode = PropVal
    PropertyChanged "OLEDropMode"
End Property

Public Property Get CellFocused(ByVal lRow As Long, _
                                ByVal lCell As Long) As Boolean
Attribute CellFocused.VB_MemberFlags = "40"
'/* spanned cell focus test
    If RowFocused(lRow) Then
        If (lCell = CellInFocus) Then
            CellFocused = True
        End If
    Else
        CellFocused = False
    End If
End Property

Public Property Let CellFocused(ByVal lRow As Long, _
                                ByVal lCell As Long, _
                                ByVal bFocus As Boolean)
'/* spanned cell focus test
    If bFocus Then
        m_lRowFocused = lRow
        CellInFocus = lCell
    Else
        CellInFocus = 0
    End If
End Property

Public Property Get RowCount() As Long
Attribute RowCount.VB_MemberFlags = "40"
'/* [get] row count
    RowCount = m_lRowCount
End Property

Public Property Get RowFocused(ByVal lRow As Long) As Boolean
Attribute RowFocused.VB_MemberFlags = "40"
'/* [get] return row focused state
    If (lRow = RowInFocus) Then
        RowFocused = True
    End If
End Property

Public Property Let RowFocused(ByVal lRow As Long, _
                               ByVal bFocus As Boolean)
'*/ [let] change row focused state
    If (lRow < m_lRowCount) Then
        If bFocus Then
            If m_bUseSpannedRows Then
                m_lRowFocused = RowSpanMapVirtual(lRow)
            Else
                RowInFocus = lRow
            End If
        Else
            RowInFocus = 0
        End If
    End If
End Property

Public Property Get RowHeight() As Long
Attribute RowHeight.VB_MemberFlags = "40"
'*/ [get] row height
    RowHeight = m_lRowHeight
End Property

Public Property Let RowHeight(ByVal PropVal As Long)
'*/ [let] row height
    If Not (PropVal = m_lRowHeight) Then
        m_lRowHeight = PropVal
        RowHeightChange
    End If
End Property

Public Property Get RowInFocus() As Long
Attribute RowInFocus.VB_MemberFlags = "40"
'*/ [get] focused row
    RowInFocus = m_lRowFocused
End Property

Public Property Let RowInFocus(ByVal lRow As Long)
'*/ [let] focused row
    m_lRowFocused = lRow
End Property

Public Property Get RowGhosted(ByVal lRow As Long) As Boolean
Attribute RowGhosted.VB_MemberFlags = "40"
'*/ [get] return row ghosted state
    If Not (m_lHGHwnd = 0) Then
        If (lRow < m_lRowCount) Then
            RowGhosted = (SendMessageLongA(m_lHGHwnd, LVM_GETITEMSTATE, lRow, LVIS_CUT))
        End If
    End If
End Property

Public Property Let RowGhosted(ByVal lRow As Long, _
                               ByVal bGhosted As Boolean)

'*/ [let] change row ghosted state

Dim uLVI As LVITEM

    If Not (m_lHGHwnd = 0) Then
        If (lRow < m_lRowCount) Then
            With uLVI
                .stateMask = LVIS_CUT
                .State = LVIS_CUT * -bGhosted
                .Mask = LVIF_STATE
            End With
            SendMessageA m_lHGHwnd, LVM_SETITEMSTATE, lRow, uLVI
        End If
    End If

End Property

Public Property Get RowSelected(ByVal lRow As Long) As Boolean
Attribute RowSelected.VB_MemberFlags = "40"
'*/ [get] return selected state
    If Not (m_lHGHwnd = 0) Then
        If (lRow < m_lRowCount) Then
            RowSelected = CBool(SendMessageLongA(m_lHGHwnd, LVM_GETITEMSTATE, lRow, LVIS_SELECTED))
        End If
    End If
End Property

Public Property Let RowSelected(ByVal lRow As Long, _
                                ByVal bSelected As Boolean)

'*/ [let] select a row

Dim uLVI As LVITEM

    If Not (m_lHGHwnd = 0) Then
        If (lRow < m_lRowCount) Then
            With uLVI
                .stateMask = LVIS_SELECTED Or -(bSelected And lRow > -1) * LVIS_FOCUSED
                .State = -bSelected * LVIS_SELECTED Or -(lRow > -1) * LVIS_FOCUSED
                .Mask = LVIF_STATE
            End With
            SendMessageA m_lHGHwnd, LVM_SETITEMSTATE, lRow, uLVI
        End If
    End If

End Property

Public Property Get RowTag(ByVal lRow As Long) As String
Attribute RowTag.VB_MemberFlags = "40"
'/* [get] row tag

    If lRow < (UBound(m_cGridItem) + 1) Then
        If lRow > (LBound(m_cGridItem) - 1) Then
            RowTag = m_cGridItem(lRow).RowTag
        End If
    End If

End Property

Public Property Let RowTag(ByVal lRow As Long, ByVal sTag As String)
'/* [let] row tag

    If lRow < (UBound(m_cGridItem) + 1) Then
        If lRow > (LBound(m_cGridItem) - 1) Then
            m_cGridItem(lRow).RowTag = sTag
        End If
    End If

End Property

Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_MemberFlags = "40"
'*/ [get] scale height
    ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Get ScaleMode() As ScaleModeConstants
Attribute ScaleMode.VB_MemberFlags = "40"
'*/ [get] scale mode
    ScaleMode = UserControl.ScaleMode
End Property

Public Property Let ScaleMode(ByVal eMode As ScaleModeConstants)
'*/ [let] scale mode
    UserControl.ScaleMode = eMode
    PropertyChanged "ScaleMode"
End Property

Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_MemberFlags = "40"
'*/ [get] scale width
    ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Get ScrollBarAlignment() As ESCScrollBarAlignment
Attribute ScrollBarAlignment.VB_MemberFlags = "40"
'/* [get] scrollbar alignment
    ScrollBarAlignment = m_eScrollBarAlignment
End Property

Public Property Let ScrollBarAlignment(ByVal PropVal As ESCScrollBarAlignment)
'/* [let] scrollbar alignment
    If Not (m_lHGHwnd = 0) Then
        If (PropVal = escRightAlign) Then
            WindowStyle m_lHGHwnd, GWL_EXSTYLE, 0, WS_EX_LEFTSCROLLBAR
        Else
            WindowStyle m_lHGHwnd, GWL_EXSTYLE, WS_EX_LEFTSCROLLBAR, 0
        End If
        Resize
    End If
    m_eScrollBarAlignment = PropVal
End Property

Public Property Get SelectedCount() As Long
Attribute SelectedCount.VB_MemberFlags = "40"
'*/ [get] retrieve selected count
    If Not (m_lHGHwnd = 0) Then
        SelectedCount = SendMessageLongA(m_lHGHwnd, LVM_GETSELECTEDCOUNT, 0&, 0&)
    End If
End Property

Public Property Get SortType() As ESTSortType
'/* [get] sort type
    SortType = m_eSortType
End Property

Public Property Let SortType(ByVal PropVal As ESTSortType)
'/* [set] sort type
    If (PropVal = estNone) Then
        HeaderFlat = True
    Else
        HeaderFlat = False
    End If
    m_eSortType = PropVal
    PropertyChanged "SortType"
End Property

Public Property Get StructPtr() As Long
Attribute StructPtr.VB_MemberFlags = "40"
'*/ [get] retrieve pointer to the data structure
    StructPtr = m_lStrctPtr
End Property

Public Property Let StructPtr(ByVal PropVal As Long)
'*/ [let] add pointer to the data structure
    If Not (m_lStrctPtr = 0) Then
        DeAllocatePointer "a", True
    End If
    m_lStrctPtr = PropVal
End Property

Public Property Get ThemeColor() As OLE_COLOR
Attribute ThemeColor.VB_MemberFlags = "40"
'/* [get] theme color
    ThemeColor = m_oThemeColor
End Property

Public Property Let ThemeColor(ByVal PropVal As OLE_COLOR)
'/* [set] theme color
    m_oThemeColor = PropVal
End Property

Public Property Get ThemeLuminence() As ESTThemeLuminence
Attribute ThemeLuminence.VB_MemberFlags = "40"
'/* [get] theme luminence
    ThemeLuminence = m_eThemeLuminence
End Property

Public Property Let ThemeLuminence(ByVal PropVal As ESTThemeLuminence)
'/* [set] theme luminence
    Select Case PropVal
    Case 0
        m_sngLuminence = 0.2
    Case 1
        m_sngLuminence = 0.4
    Case 2
        m_sngLuminence = 0.7
    End Select
    m_eThemeLuminence = PropVal
End Property

Public Property Get UseThemeColors() As Boolean
Attribute UseThemeColors.VB_MemberFlags = "40"
'/* [get] theme status
    UseThemeColors = m_bUseThemeColors
End Property

Public Property Let UseThemeColors(ByVal PropVal As Boolean)
'/* [set] theme option
    m_bUseThemeColors = PropVal
    m_bUseCheckBoxTheme = PropVal
End Property

Public Property Get UseUnicode() As Boolean
Attribute UseUnicode.VB_MemberFlags = "40"
'/* [get] unicode state
    UseUnicode = m_bUseUnicode
End Property

Public Property Let UseUnicode(ByVal PropVal As Boolean)
'/* [let] unicode state
    If m_bIsNt Then
        If PropVal Then
            m_bUseUnicode = True
        Else
            m_bUseUnicode = False
        End If
        If Not (m_lHGHwnd = 0) Then
            SetUnicode m_bUseUnicode
        End If
    End If
End Property

Public Property Get VirtualMode() As Boolean
'/* [get] virtual mode
    VirtualMode = m_bVirtualMode
End Property

Public Property Let VirtualMode(ByVal PropVal As Boolean)
'/* [let] virtual mode
    m_bVirtualMode = PropVal
    PropertyChanged "VirtualMode"
End Property

Public Property Get XPColors() As Boolean
'/* [get] use xp colors
    XPColors = m_bXPColors
End Property

Public Property Let XPColors(ByVal PropVal As Boolean)
'/* [let] use xp colors
    m_bXPColors = PropVal
    PropertyChanged "XPColors"
End Property


'**********************************************************************
'*                              SUPPORT
'**********************************************************************

Public Sub CheckAll()
'*/ mark all checkboxes

Dim lCt As Long

    If Not (m_lHGHwnd = 0) Then
        If ArrayCheck(m_cGridItem) Then
            For lCt = LBound(m_cGridItem) To UBound(m_cGridItem)
                m_cGridItem(lCt).Checked = True
            Next lCt
            GridRefresh False
        End If
    End If

End Sub

Private Sub CheckBoxMetrics()
'/* checkbox system metrics

    m_lCheckWidth = GetSystemMetrics(SM_CXSMICON)
    m_lCheckHeight = GetSystemMetrics(SM_CYSMICON)
    If (m_lCheckWidth = 0) Or (m_lCheckHeight = 0) Then
        m_lCheckWidth = 16
        m_lCheckHeight = 16
    End If
    
End Sub

Private Function CheckToggle(ByVal lRow As Long) As Boolean
'/* toggle check state

    If ArrayCheck(m_cGridItem) Then
        m_cGridItem(lRow).Checked = Not m_cGridItem(lRow).Checked
    End If
    RaiseEvent eVHItemCheck(lRow, CheckToggle)
    
End Function

Public Function ColumnAdd(ByVal lColumn As Long, _
                          ByVal sText As String, _
                          ByVal lWidth As Long, _
                          Optional ByVal eAlign As ECAColumnAlign = ecaColumnLeft, _
                          Optional ByVal lIcon As Long = -1, _
                          Optional ByVal eColumnTag As ECSColumnSortTags = ecsSortAuto)

'*/ create column headers

Dim bFirst As Boolean
Dim uLVC   As LVCOLUMN
Dim uHDI   As HDITEM
Dim uHDW   As HDITEMW

    If Not (m_lHGHwnd = 0) Then
        m_lHdrHwnd = HeaderHwnd()
        If Not (m_lHdrHwnd = 0) Then
            ReDim Preserve m_bColumnLock(0 To lColumn)
            bFirst = (ColumnCount = 0)
            '/* sort flag
            If c_ColumnTags Is Nothing Then
                Set c_ColumnTags = New Collection
            End If
            On Error Resume Next
            c_ColumnTags.Add eColumnTag, CStr(lColumn)
            On Error GoTo 0
            '/* unicode nt
            With uLVC
                .pszText = StrPtr(sText)
                .cchTextMax = LenB(sText)
                .cx = lWidth
                .fmt = eAlign
                .Mask = LVCF_TEXT Or LVCF_WIDTH Or LVCF_FMT
            End With
            '/* add the column
            If m_bIsNt Then
                ColumnAdd = (SendMessageW(m_lHGHwnd, LVM_INSERTCOLUMNW, lColumn, uLVC) > -1)
            '/* no unicode
            Else
                ColumnAdd = (SendMessageA(m_lHGHwnd, LVM_INSERTCOLUMNA, lColumn, uLVC) > -1)
            End If
            '/* add the imagelist
            If Not (ColumnAdd = 0) Then
                If bFirst Then
                    If Not (m_lImlHdHndl = 0) Then
                        If m_bIsNt Then
                            SendMessageLongW m_lHdrHwnd, HDM_SETIMAGELIST, 0&, m_lImlHdHndl
                        Else
                            SendMessageLongA m_lHdrHwnd, HDM_SETIMAGELIST, 0&, m_lImlHdHndl
                        End If
                    End If
                End If
                '/* unicode nt
                If m_bIsNt Then
                    '/* build data struct
                    With uHDW
                        .pszText = StrPtr(sText)
                        .cchTextMax = LenB(sText)
                        .cxy = lWidth
                        If Not (m_lImlHdHndl = 0) Then
                            .iImage = lIcon
                        Else
                            lIcon = -1
                        End If
                        .fmt = HDF_STRING Or eAlign * -(lColumn <> 0) Or HDF_IMAGE * -(lIcon > -1) Or HDF_BITMAP_ON_RIGHT
                        .Mask = HDI_TEXT Or HDI_WIDTH Or HDI_FORMAT Or HDI_IMAGE * -(lIcon > -1)
                    End With
                    '/* pass to header
                    SendMessageW m_lHdrHwnd, HDM_SETITEMW, lColumn, uHDW
                Else
                    With uHDI
                        .pszText = sText
                        .cchTextMax = LenB(sText)
                        .cxy = lWidth
                        If Not (m_lImlHdHndl = 0) Then
                            .iImage = lIcon
                        Else
                            lIcon = -1
                        End If
                        .fmt = HDF_STRING Or eAlign * -(lColumn <> 0) Or HDF_IMAGE * -(lIcon > -1) Or HDF_BITMAP_ON_RIGHT
                        .Mask = HDI_TEXT Or HDI_WIDTH Or HDI_FORMAT Or HDI_IMAGE * -(lIcon > -1)
                    End With
                    SendMessageA m_lHdrHwnd, HDM_SETITEMA, lColumn, uHDI
                End If
            End If
            '/* track column count
            If Not (m_cSkinHeader Is Nothing) Then
                m_cSkinHeader.ColumnCountChange = lColumn
            End If
            If (m_lColumnHeight = 0) Then
                m_lColumnHeight = ColumnHeight
            End If
            m_lColumnCount = ColumnCount
            RaiseEvent eVHColumnAdded(lColumn, lWidth, lIcon, sText)
        End If
    End If
    
End Function

Public Sub ColumnAutosize(ByVal lColumn As Long, _
                          Optional ByVal AutosizeType As ECUColumnAutosize = ecuColumnItem)

'*/ autosize columns
    If Not (m_lHGHwnd = 0) Then
        If (lColumn < m_lColumnCount) Then
            SendMessageLongA m_lHGHwnd, LVM_SETCOLUMNWIDTH, lColumn, AutosizeType
        End If
    End If
End Sub

Public Function ColumnClear() As Boolean
'*/ remove all columns

Dim lCt As Long

    For lCt = (m_lColumnCount - 1) To 0 Step -1
        ColumnRemove lCt
    Next lCt
    
End Function

Private Sub ColumnIconReset()
'/* reset column icons

Dim lCt As Long

    For lCt = 0 To (m_lColumnCount - 1)
        ColumnIcon(lCt) = -1
    Next lCt

End Sub

Public Function ColumnLastFit() As Boolean
'/* fit last column

Dim lCol    As Long

    If Not (m_lHGHwnd = 0) Then
        lCol = (m_lColumnCount - 1)
        If m_bIsNt Then
            SendMessageLongW m_lHGHwnd, LVM_SETCOLUMNWIDTH, lCol, LVSCW_AUTOSIZE_USEHEADER
        Else
            SendMessageLongA m_lHGHwnd, LVM_SETCOLUMNWIDTH, lCol, LVSCW_AUTOSIZE_USEHEADER
        End If
    End If

End Function

Public Function ColumnRemove(ByVal lColumn As Long) As Boolean
'*/ remove a column

    If Not (m_lHGHwnd = 0) Then
        If (lColumn < m_lColumnCount) Then
            If Not m_bVirtualMode Then
                If RowCount > 0 Then
                    CellRemoveColumn lColumn
                End If
            End If
            If m_bIsNt Then
                ColumnRemove = CBool(SendMessageLongW(m_lHGHwnd, LVM_DELETECOLUMN, lColumn, 0&))
            Else
                ColumnRemove = CBool(SendMessageLongA(m_lHGHwnd, LVM_DELETECOLUMN, lColumn, 0&))
            End If
            On Error Resume Next
            c_ColumnTags.Remove CStr(lColumn)
            On Error GoTo 0
            RaiseEvent eVHColumnRemoved(lColumn)
        End If
    End If
    m_lColumnCount = ColumnCount

End Function

Private Sub CellRemoveColumn(ByVal lColumn As Long)
'/* remove column cells

Dim lCt As Long
Dim lUb As Long
Dim lLb As Long

    If (lColumn < m_lColumnCount) Then
        If ArrayCheck(m_cGridItem) Then
            lLb = LBound(m_cGridItem)
            lUb = UBound(m_cGridItem)
            Do
                m_cGridItem(lCt).RemoveCell lColumn
                lCt = lCt + 1
            Loop Until lCt > lUb
        End If
    End If

End Sub

Private Sub ColumnSizeHeight()
'/* column fit text height

Dim tPnt As POINTAPI

    If LeftKeyState Then
        GetCursorPos tPnt
        ScreenToClient UserControl.hwnd, tPnt
        With tPnt
            If Not (.y = m_lHeaderHeight) Then
                If (Abs(.y - m_lHeaderHeight) > 2) Then
                    HeaderHeight = .y
                    RaiseEvent eVHColumnVerticalSize(.y)
                End If
            End If
        End With
        m_bTransitionMask = True
    Else
        m_bTransitionMask = False
    End If

End Sub

Public Sub ColumnSizeToItems(Optional ByVal bColumnFit As Boolean)
'/* size columns to longest items

Dim lCol   As Long
Dim lParam As Long

    If Not (m_lHGHwnd = 0) Then
        If bColumnFit Then
            lParam = LVSCW_AUTOSIZE_USEHEADER
        Else
            lParam = LVSCW_AUTOSIZE
        End If
        '/* size all columns
        For lCol = 0 To (m_lColumnCount - 1)
            If m_bIsNt Then
                SendMessageLongW m_lHGHwnd, LVM_SETCOLUMNWIDTH, lCol, lParam
            Else
                SendMessageLongA m_lHGHwnd, LVM_SETCOLUMNWIDTH, lCol, lParam
            End If
        Next lCol
    End If

End Sub

Public Function ColumnTextBestFit(ByVal lColumn As Long) As Long
'/* fit row heights to column text

Dim lHdc    As Long
Dim lSt     As Long
Dim lLt     As Long
Dim lMt     As Long
Dim lCw     As Long
Dim lVt     As Long
Dim lRCt    As Long
Dim tPnt    As POINTAPI

    If Not m_bVirtualMode Then
        lSt = FindShortestString(lColumn)
        lLt = FindLongestString(lColumn)
        lMt = (lSt \ 2) + (lLt \ 2)
        lCw = ColumnWidth(lColumn)
        lHdc = GetDC(m_lHGHwnd)
        CharDimensions lHdc, tPnt
        ReleaseDC m_lHGHwnd, lHdc
        With tPnt
            lRCt = (.x * lMt) \ lCw
            lVt = (lRCt * .y) / 2
        End With
        If lVt > m_lRowMinHgt Then
            ColumnTextBestFit = lVt
            m_lRowHeight = lVt
            RowHeightChange
        Else
            ColumnTextBestFit = m_lRowMinHgt
        End If
    End If

End Function

Public Function ColumnTextFitHeight(ByVal lColumn As Long) As Long
'/* column cells text best fit
Dim lHdc    As Long
Dim lLt     As Long
Dim lCw     As Long
Dim lVt     As Long
Dim lRCt    As Long
Dim tPnt    As POINTAPI

    If (lColumn < m_lColumnCount) Then
        If Not m_bVirtualMode Then
            lLt = FindLongestString(lColumn)
            lCw = ColumnWidth(lColumn)
            lHdc = GetDC(m_lHGHwnd)
            CharDimensions lHdc, tPnt
            ReleaseDC m_lHGHwnd, lHdc
            With tPnt
                lRCt = (.x * lLt) \ lCw
                lVt = (lRCt * .y) / 2
            End With
            If (lVt > m_lRowMinHgt) Then
                ColumnTextFitHeight = lVt
                m_lRowHeight = lVt
                RowHeightChange
            Else
                ColumnTextFitHeight = m_lRowMinHgt
            End If
        End If
    End If

End Function

Public Sub CopyCellToClipboard()
'/* copy selected item to clipboard

Dim sTemp As String

On Error GoTo Handler

    sTemp = CellText(RowInFocus, CellInFocus)
    If (Len(sTemp) > 0) Then
        With Clipboard
            .Clear
            .SetText sTemp
        End With
    End If

Handler:
    On Error GoTo 0
    
End Sub

Private Function DividerPosition() As Long
'/* get drag divider coords

Dim lPos  As Long
Dim tPnt  As POINTAPI
Dim tRect As RECT

    If Not (m_lHdrHwnd = 0) Then
        GetCursorPos tPnt
        ScreenToClient m_lHdrHwnd, tPnt
        GetClientRect m_lHdrHwnd, tRect
        With tPnt
            If (.y > -8) And (.y < tRect.Bottom + 8) Then
                lPos = (.x And &HFFFF&)
                lPos = lPos Or (.y And &H7FFF) * &H10000
                If (.y And &H8000) = &H8000 Then
                    lPos = lPos Or &H80000000
                End If
                DividerPosition = SendMessageLongA(m_lHdrHwnd, HDM_SETHOTDIVIDER, 1&, lPos)
            Else
                DividerPosition = -1
            End If
        End With
    End If

End Function

Private Function DragStartTimer() As Boolean
'/* start header drag timer

    If Not m_bTimerActive Then
        If Not (m_lParentHwnd = 0) Then
            SetTimer m_lParentHwnd, 1&, 10&, 0&
            m_bTimerActive = True
        End If
    End If

End Function

Private Function DragStopTimer() As Boolean
'/* stop header drag timer

    If m_bTimerActive Then
        If Not (m_lParentHwnd = 0) Then
            KillTimer m_lParentHwnd, 1&
            m_bTimerActive = False
            m_lSafeTimer = 0
        End If
    End If

End Function

Private Function FontHandle(ByVal oFont As StdFont) As Long
'*/ change font

Dim lChar       As Long
Dim uLF         As LOGFONT
Dim bteFont()   As Byte

    If Not (oFont Is Nothing) Then
        With uLF
            bteFont = StrConv(oFont.Name, vbFromUnicode)
            For lChar = 0 To UBound(bteFont)
                .lfFaceName(lChar) = bteFont(lChar)
            Next lChar
            .lfHeight = -MulDiv(oFont.Size, GetDeviceCaps(UserControl.hdc, LOGPIXELSY), 72)
            .lfItalic = oFont.Italic
            .lfWeight = IIf(oFont.Bold, FW_BOLD, FW_NORMAL)
            .lfUnderline = oFont.Underline
            .lfStrikeOut = oFont.Strikethrough
            .lfCharSet = oFont.Charset
            If m_bIsXp Then
                .lfQuality = LF_CLEARTYPE_QUALITY
            Else
                .lfQuality = LF_ANTIALIASED_QUALITY
            End If
        End With
        If m_bUseUnicode Then
            FontHandle = CreateFontIndirectW(uLF)
        Else
            FontHandle = CreateFontIndirectA(uLF)
        End If
    End If

End Function

Private Function RowSpanMapVirtual(ByVal lRow As Long) As Long

Dim lCt     As Long
Dim lOffset As Long

    lOffset = (m_lRowHeight * lRow)
    For lCt = LBound(m_cGridItem) To UBound(m_cGridItem)
        If (m_lRowDepth(lCt) > lOffset) Then
            RowSpanMapVirtual = lCt
            Exit For
        End If
    Next lCt

End Function

Private Function ScrollPosition() As Long
'/* current dc Y postion
    ScrollPosition = (RowTopIndex * m_lRowHeight)
End Function

Public Sub ReturnCellRect(ByVal lRow As Long, _
                          ByVal lCell As Long, _
                          ByRef lLeft As Long, _
                          ByRef lTop As Long, _
                          ByRef lRight As Long, _
                          ByRef lBottom As Long)

Dim tRect As RECT

    GetCellRect lRow, lCell, tRect
    With tRect
        lLeft = .left
        lTop = .top
        lRight = .Right
        lBottom = .Bottom
    End With
    
End Sub

Private Sub GetCellRect(ByVal lRow As Long, _
                        ByVal lCell As Long, _
                        ByRef tRect As RECT)
'/* get rect struct of row cell

Dim lColumn As Long
Dim tRfc    As RECT
Dim tRtop   As RECT

    If Not (m_lHGHwnd = 0) Then
        With tRect
            .left = LVIR_LABEL
            .top = lCell
            If m_bIsNt Then
                SendMessageW m_lHGHwnd, LVM_GETSUBITEMRECT, lRow, tRect
            Else
                SendMessageA m_lHGHwnd, LVM_GETSUBITEMRECT, lRow, tRect
            End If
            If m_bUseSpannedRows Then
                If (lRow = 0) Then
                    .Bottom = .top + m_lRowDepth(lRow)
                Else
                    tRtop.left = LVIR_BOUNDS
                    If m_bIsNt Then
                        SendMessageW m_lHGHwnd, LVM_GETITEMRECT, RowTopIndex, tRtop
                    Else
                        SendMessageA m_lHGHwnd, LVM_GETITEMRECT, RowTopIndex, tRtop
                    End If
                    .top = tRtop.top + (m_lRowDepth(lRow - 1) - ScrollPosition)
                    .Bottom = .top + (m_lRowDepth(lRow) - m_lRowDepth(lRow - 1))
                End If
            End If
            If (lCell = 0) Then
                lColumn = ColumnIndex(lCell)
                If Not (lColumn = 0) Then
                    lColumn = ColumnAtIndex(lColumn - 1)
                    With tRfc
                        .left = LVIR_LABEL
                        .top = lColumn
                    End With
                    If m_bIsNt Then
                        SendMessageW m_lHGHwnd, LVM_GETSUBITEMRECT, lRow, tRfc
                    Else
                        SendMessageA m_lHGHwnd, LVM_GETSUBITEMRECT, lRow, tRfc
                    End If
                    .left = tRfc.Right
                Else
                    .left = (.Right - ColumnWidth(0))
                End If
            End If
        End With
    End If

End Sub

Private Sub GetRowRect(ByVal lRow As Long, _
                       ByRef tRect As RECT)

'/* get rect struct of row item

Dim tRcpy As RECT

    If Not (m_lHGHwnd = 0) Then
        If Not (lRow = -1) Then
            With tRect
                .left = LVIR_BOUNDS
                SendMessageA m_lHGHwnd, LVM_GETITEMRECT, lRow, tRect
                GetCellRect lRow, 0, tRcpy
                .top = tRcpy.top
                .Bottom = tRcpy.Bottom
            End With
        End If
    End If

End Sub

Private Function GetMaskColor(ByVal lHdc As Long) As Long
'/* get checkbox mask color
    GetMaskColor = GetPixel(lHdc, 0&, 0&)
End Function

Private Sub GridLine(ByVal lHdc As Long, _
                     ByVal X1 As Long, _
                     ByVal y1 As Long, _
                     ByVal x2 As Long, _
                     ByVal y2 As Long, _
                     ByVal lColor As Long, _
                     ByVal lWidth As Long)

'/* draw gridline

Dim lhPen       As Long
Dim lhPenOld    As Long
Dim tPnt          As POINTAPI

    lhPen = CreatePen(0, lWidth, lColor)
    lhPenOld = SelectObject(lHdc, lhPen)
    MoveToEx lHdc, X1, y1, tPnt
    LineTo lHdc, x2, y2
    SelectObject lHdc, lhPenOld
    DeleteObject lhPen

End Sub

Public Sub GridRefresh(Optional ByVal bErase As Boolean)
'/* refresh the listview

Dim tRect As RECT

    If Not (m_lHGHwnd = 0) Then
        GetClientRect m_lHGHwnd, tRect
        If bErase Then
            EraseRect m_lHGHwnd, tRect, 1&
        Else
            EraseRect m_lHGHwnd, tRect, 0&
        End If
    End If

End Sub

Private Function HotDivider() As Long
'/* draw divider insertion mark

Dim lDivPos     As Long
Dim lPrHnd      As Long
Dim lParDc      As Long
Dim lGrdDc      As Long
Dim lCol        As Long
Dim lColIdx     As Long
Dim lCount      As Long
Dim lXPos       As Long
Dim lhPen       As Long
Dim lhPenOld    As Long
Dim tPnt        As POINTAPI
Dim tPcd        As POINTAPI
Dim tPos        As POINTAPI
Dim tRect       As RECT
Dim tRClt       As RECT
Dim tRHdr       As RECT

    '/* test boundaries
    GetCursorPos tPos
    ScreenToClient m_lHdrHwnd, tPos
    GetClientRect m_lHGHwnd, tRClt
    lPrHnd = GetParent(m_lParentHwnd)
    If (lPrHnd = 0) Then Exit Function

    '/* column divider index
    lDivPos = DividerPosition
    lCount = m_lColumnCount
    '/* scrolled offscreen
    If (lDivPos > tRClt.Right) Then
        Exit Function
    End If
    If (lDivPos = -1) Then
        Exit Function
    End If
    '/* 20 second timeout
    If (m_lSafeTimer > 2000) Then
        DragStopTimer
        Exit Function
    Else
        m_lSafeTimer = m_lSafeTimer + 1
    End If
    '/* relative column
    If lDivPos = lCount Then
        lCol = lCount - 1
    Else
        lCol = lDivPos
    End If

    lColIdx = m_cSkinHeader.ColumnAtIndex(lCol)
    SendMessageA m_lHdrHwnd, HDM_GETITEMRECT, lColIdx, tRect
    GetWindowRect m_lHdrHwnd, tRHdr
    CopyMemory tPcd, tRHdr, Len(tPcd)
    ScreenToClient m_lHGHwnd, tPcd

    '/* scrolled off control
    If (tPos.x > tRect.Right) Then
        If (m_eDragEffectStyle = edsClientArrow) Then
            HotDividerReset
            Exit Function
        End If
    End If
    
    OffsetRect tRect, tPcd.x, 0
    '/* position mark
    If lDivPos = 0 Then
        lXPos = 0
    ElseIf lDivPos = lCount Then
        lXPos = tRect.Right
    Else
        lXPos = tRect.left
    End If
    
    If (m_eDragEffectStyle = edsClientArrow) Then
        '/* refresh
        If Not (m_tDividerRect(0).left = (lXPos - 5)) Then
            HotDividerReset
        End If
        lGrdDc = GetDC(m_lHGHwnd)
        With tRect
            '/* bottom arrow
            lhPen = CreatePen(0&, 2&, 255&)
            lhPenOld = SelectObject(lGrdDc, lhPen)
            MoveToEx lGrdDc, lXPos, (.Bottom + 11), tPnt
            LineTo lGrdDc, lXPos, (.Bottom + 6)
            SelectObject lGrdDc, lhPenOld
            DeleteObject lhPen
            lhPen = CreatePen(0&, 1&, 255&)
            lhPenOld = SelectObject(lGrdDc, lhPen)
            MoveToEx lGrdDc, (lXPos - 5), (.Bottom + 5), tPnt
            LineTo lGrdDc, (lXPos + 5), (.Bottom + 5)
            MoveToEx lGrdDc, (lXPos - 4), (.Bottom + 4), tPnt
            LineTo lGrdDc, (lXPos + 4), (.Bottom + 4)
            MoveToEx lGrdDc, (lXPos - 3), (.Bottom + 3), tPnt
            LineTo lGrdDc, (lXPos + 3), (.Bottom + 3)
            MoveToEx lGrdDc, (lXPos - 2), (.Bottom + 2), tPnt
            LineTo lGrdDc, (lXPos + 2), (.Bottom + 2)
            MoveToEx lGrdDc, (lXPos - 1), (.Bottom + 1), tPnt
            LineTo lGrdDc, (lXPos + 1), (.Bottom + 1)
            SelectObject lGrdDc, lhPenOld
            DeleteObject lhPen
        End With
        '/* store coords
        With m_tDividerRect(0)
            .Bottom = tRect.Bottom + 12
            .left = lXPos - 5
            .Right = lXPos + 5
            .top = tRect.Bottom - 1
        End With
        '/* cleanup
        ReleaseDC m_lHGHwnd, lGrdDc
        lhPen = 0
        lhPenOld = 0
    
        '/* top arrow
        lParDc = GetDC(lPrHnd)
        GetWindowRect m_lHGHwnd, tRect
        CopyMemory tPcd, tRect, Len(tPcd)
        ScreenToClient lPrHnd, tPcd

        With tPcd
            lXPos = (.x + lXPos)
            lhPen = CreatePen(0&, 2&, 255&)
            lhPenOld = SelectObject(lParDc, lhPen)
            MoveToEx lParDc, lXPos, (.y - 12), tPnt
            LineTo lParDc, lXPos, (.y - 7)
            SelectObject lParDc, lhPenOld
            DeleteObject lhPen
            lhPen = CreatePen(0&, 1&, 255&)
            lhPenOld = SelectObject(lParDc, lhPen)
            MoveToEx lParDc, (lXPos - 5), (.y - 6), tPnt
            LineTo lParDc, (lXPos + 5), (.y - 6)
            MoveToEx lParDc, (lXPos - 4), (.y - 5), tPnt
            LineTo lParDc, (lXPos + 4), (.y - 5)
            MoveToEx lParDc, (lXPos - 3), (.y - 4), tPnt
            LineTo lParDc, (lXPos + 3), (.y - 4)
            MoveToEx lParDc, (lXPos - 2), (.y - 3), tPnt
            LineTo lParDc, (lXPos + 2), (.y - 3)
            MoveToEx lParDc, (lXPos - 1), (.y - 2), tPnt
            LineTo lParDc, (lXPos + 1), (.y - 2)
            SelectObject lGrdDc, lhPenOld
            DeleteObject lhPen
        End With
        '/* store parent rect
        With m_tDividerRect(1)
            .Bottom = tPcd.y - 1
            .left = lXPos - 5
            .Right = lXPos + 5
            .top = tPcd.y - 12
        End With
        ReleaseDC lPrHnd, lParDc
    Else
        lGrdDc = GetDC(m_lHdrHwnd)
        If (m_eDragEffectStyle = edsThinLine) Then
            lhPen = CreatePen(0&, 1&, &H808080)
        Else
            lhPen = CreatePen(0&, 2&, &H808080)
        End If
        lhPenOld = SelectObject(lGrdDc, lhPen)
        With tRect
            MoveToEx lGrdDc, (lXPos - 1), (.top - 1), tPnt
            LineTo lGrdDc, (lXPos - 1), (.Bottom - 1)
        End With
        SelectObject lGrdDc, lhPenOld
        DeleteObject lhPen
        lhPenOld = 0
        lhPen = 0
        ReleaseDC m_lHdrHwnd, lGrdDc
    End If

End Function

Private Sub HotDividerReset()
'/* erase divider mark

Dim lPrHnd  As Long
Dim tRect   As RECT

    '/* list arrow
    EraseRect m_lHGHwnd, m_tDividerRect(0), 0&
    '/* parent arrow
    lPrHnd = GetParent(m_lParentHwnd)
    If lPrHnd = 0 Then Exit Sub
    EraseRect lPrHnd, m_tDividerRect(1), 0&
    '/* non client
    GetClientRect m_lHdrHwnd, tRect
    With tRect
        .top = .Bottom + 1
        .Bottom = .top + 2
    End With
    EraseRect m_lHGHwnd, tRect, 1&


End Sub

Public Sub ImlHeaderAddBmp(ByVal lBitmap As Long, _
                           Optional ByVal lMaskColor As Long = CLR_NONE)

'*/ add a bitmap to header iml

On Error GoTo Handler

    If Not (m_lImlHdHndl = 0) Then
        If Not (lMaskColor = CLR_NONE) Then
            ImageList_AddMasked m_lImlHdHndl, lBitmap, lMaskColor
        Else
            ImageList_Add m_lImlHdHndl, lBitmap, 0&
        End If
    End If

Handler:
    On Error GoTo 0

End Sub

Public Sub ImlHeaderAddIcon(ByVal lIcon As Long)
'*/ add an icon to header iml

On Error GoTo Handler

    If Not (m_lImlHdHndl = 0) Then
        ImageList_AddIcon m_lImlHdHndl, lIcon
    End If

Handler:
    On Error GoTo 0

End Sub

Public Sub ImlRowAddBmp(ByVal lBitmap As Long, _
                        Optional ByVal lMaskColor As Long = CLR_NONE)

'*/ add bmp to small image iml

On Error GoTo Handler

    If Not (m_lImlRowHndl = 0) Then
        If Not (lMaskColor = CLR_NONE) Then
            ImageList_AddMasked m_lImlRowHndl, lBitmap, lMaskColor
        Else
            ImageList_Add m_lImlRowHndl, lBitmap, 0&
        End If
    End If

Handler:
    On Error GoTo 0

End Sub

Public Sub ImlRowAddIcon(ByVal lIcon As Long)
'*/ add icon to small image iml

On Error GoTo Handler

    If Not (m_lImlRowHndl = 0) Then
        ImageList_AddIcon m_lImlRowHndl, lIcon
    End If

Handler:
    On Error GoTo 0

End Sub

Public Sub ImlStateAddBmp(ByVal lBitmap As Long, _
                          Optional ByVal lMaskColor As Long = CLR_NONE)

'*/ add a bitmap to header iml

On Error GoTo Handler

    If Not (m_lImlStateHndl = 0) Then
        If Not (lMaskColor = CLR_NONE) Then
            ImageList_AddMasked m_lImlStateHndl, lBitmap, lMaskColor
        Else
            ImageList_Add m_lImlStateHndl, lBitmap, 0&
        End If
    End If

Handler:
    On Error GoTo 0

End Sub

Public Sub ImlStateAddIcon(ByVal lIcon As Long)
'*/ add an icon to header iml

On Error GoTo Handler

    If Not (m_lImlStateHndl = 0) Then
        ImageList_AddIcon m_lImlStateHndl, lIcon
    End If

Handler:
    On Error GoTo 0

End Sub

Public Sub ImlDragAddBmp(ByVal lBitmap As Long, _
                         Optional ByVal lMaskColor As Long = CLR_NONE)

'*/ add bmp to ole drag iml

On Error GoTo Handler

    If Not (m_lDragImgIml = 0) Then
        If Not (lMaskColor = CLR_NONE) Then
            ImageList_AddMasked m_lDragImgIml, lBitmap, lMaskColor
        Else
            ImageList_Add m_lDragImgIml, lBitmap, 0&
        End If
    End If

Handler:
    On Error GoTo 0

End Sub

Public Sub InitImlHeader()
'*/ initialize header imagelist

    If Not (m_lHGHwnd = 0) Then
        m_lImlHdHndl = m_cHeaderIcons.hIml
    End If

End Sub

Public Sub InitImlRow(Optional ByVal lWidth As Long = 16, _
                      Optional ByVal lHeight As Long = 16)

'*/ initialize smallicons image list

    If Not (m_lHGHwnd = 0) Then
        m_lImlRowHndl = m_cCellIcons.hIml
        If Not (m_lImlRowHndl = 0) Then
            SendMessageLongA m_lHGHwnd, LVM_SETIMAGELIST, LVSIL_SMALL, m_lImlRowHndl
            m_lRowIconX = lWidth
            m_lRowIconY = lHeight
            m_lRowMinHgt = m_lRowIconY + 6
        End If
    End If

End Sub

Public Sub InitImlState()
'*/ initialize header imagelist

    If Not (m_lHGHwnd = 0) Then
        DestroyImlState
        m_lImlStateHndl = ImageList_Create(16&, 16&, ILC_COLOR32 Or ILC_MASK, 0&, 0&)
        If Not (m_lImlStateHndl = 0) Then
            SendMessageLongA m_lHGHwnd, LVM_SETIMAGELIST, LVSIL_STATE, m_lImlStateHndl
        End If
    End If

End Sub

Public Sub InitImlDrag(ByVal lWidth As Long, _
                       ByVal lHeight As Long)
'*/ initialize header imagelist

    If Not (m_lHGHwnd = 0) Then
        DestroyImlDrag
        m_lDragImgIml = ImageList_Create(lWidth, lHeight, ILC_MASK Or ILC_COLOR32, 1&, 1&)
    End If

End Sub

Public Sub Refresh()
'/* refresh grid

    If Not (m_lHGHwnd = 0) Then
        If m_bIsNt Then
            SendMessageLongW m_lHGHwnd, LVM_UPDATE, 0&, 0&
        Else
            SendMessageLongA m_lHGHwnd, LVM_UPDATE, 0&, 0&
        End If
    End If

End Sub

Public Function RemoveDuplicates() As Boolean
'*/ remove duplicates
'TODO

Dim lCt As Long
Dim lLb As Long
Dim lUb As Long
Dim cT  As Collection

On Error GoTo Handler

    If Not m_bVirtualMode Then
        BuildStringSortArray 0
        Set cT = New Collection
        m_bSorted = False
        '/* get bounds
        lLb = LBound(m_sSortArray)
        lUb = UBound(m_sSortArray)
        lCt = lLb
        '/* filter with collection key
        Do
            cT.Add 1, m_sSortArray(lCt)
            If (Err.Number = 457) Then
                Set m_cGridItem(lCt) = Nothing
            End If
            Err.Clear
            lCt = lCt + 1
        Loop Until (lCt > lUb)
        '/* reset the array
        ResetArray m_cGridItem
        '/* init list
        SetRowCount (UBound(m_cGridItem) + 1)

        '/* success
        m_bSorted = False
        RemoveDuplicates = True
    End If

Handler:
    On Error GoTo 0

End Function

Private Sub RePaint()
'/* repaint grid

    If Not (m_lHGHwnd = 0) Then
        SendMessageLongA m_lHGHwnd, WM_PAINT, 0&, 0&
    End If

End Sub

Public Sub Resize()
'/* resize grid

Dim lFlags  As Long
Dim lSize  As Long
Dim tRect   As RECT

On Error Resume Next

    lFlags = SWP_NOACTIVATE Or SWP_NOOWNERZORDER Or SWP_NOZORDER
    If Not (m_lHGHwnd = 0) Then
        If Not (m_lParentHwnd = 0) Then
            GetClientRect m_lParentHwnd, tRect
            With tRect
                If Not (m_cTreeView Is Nothing) Then
                    Select Case m_eTvControlAlignment
                    Case etvLeftAlign
                        lSize = m_cTreeView.Width
                        SetWindowPos m_lHGHwnd, 0&, lSize, .top, (.Right - lSize), .Bottom, lFlags
                        m_cTreeView.Resize 0
                    Case etvRightAlign
                        lSize = m_cTreeView.Width
                        SetWindowPos m_lHGHwnd, 0&, 0&, .top, (.Right - lSize) - 10, .Bottom, lFlags
                        m_cTreeView.Resize 1
                    Case etvTopAlign
                        lSize = m_cTreeView.Height
                        SetWindowPos m_lHGHwnd, 0&, 0&, lSize + 10, .Right, (.Bottom - lSize) - 10, lFlags
                        m_cTreeView.Resize 2
                    Case etvBottomAlign
                        lSize = m_cTreeView.Height
                        SetWindowPos m_lHGHwnd, 0&, 0&, 0&, .Right, (.Bottom - lSize) - 10, lFlags
                        m_cTreeView.Resize 3
                    End Select
                Else
                    SetWindowPos m_lHGHwnd, 0&, 0&, .top, .Right, .Bottom, lFlags
                End If
                If Not (m_cSkinScrollBars Is Nothing) Then
                    m_cSkinScrollBars.Resize
                    m_cSkinScrollBars.ByPassHitTest = False
                    m_cSkinScrollBars.Refresh
                End If
                If Not (m_cSkinHeader Is Nothing) Then
                    m_cSkinHeader.Refresh -1
                End If
                RaiseEvent eVHGridSizeChange(.Right, .Bottom)
            End With
        End If
    End If

On Error GoTo 0

End Sub

Private Function TreeViewDividerHitTest() As Boolean
'/* treeview divider  hittest

Dim tRect   As RECT
Dim tPnt    As POINTAPI

    If Not (m_cTreeView Is Nothing) Then
        GetCursorPos tPnt
        ScreenToClient m_lParentHwnd, tPnt
        GetClientRect m_lParentHwnd, tRect
        Select Case m_eTvControlAlignment
        Case etvLeftAlign, etvRightAlign
            With tRect
                .top = (.Bottom / 2) - 12
                .Bottom = .top + 24
            End With
        Case etvTopAlign, etvBottomAlign
            With tRect
                .left = (.Right / 2) - 12
                .Right = .left + 24
            End With
        End Select
        If Not (PtInRect(tRect, tPnt.x, tPnt.y) = 0) Then
            TreeViewDividerHitTest = True
        End If
    End If

End Function

Private Sub TreeViewSizeWidth()
'/* size treeview

Dim lSize   As Long
Dim tPnt    As POINTAPI
Dim tRect   As RECT

    m_bTreeViewSizing = False
    If m_bTreeViewSizeable Then
        If TreeViewDividerHitTest Then
            If Not (m_cSkinScrollBars Is Nothing) Then
                m_cSkinScrollBars.ByPassHitTest = True
            End If
            GetCursorPos tPnt
            ScreenToClient m_lParentHwnd, tPnt
            If LeftKeyState Then
                If (m_eTvControlAlignment < 2) Then
                    lSize = tPnt.x
                    If (lSize > 10) Then
                        GetClientRect m_lParentHwnd, tRect
                        If (lSize < (tRect.Right - 10)) Then
                            m_bTreeViewSizing = True
                            If (m_eTvControlAlignment = etvLeftAlign) Then
                                m_cTreeView.Width = lSize
                            Else
                                m_cTreeView.Width = tRect.Right - lSize
                            End If
                            If (Abs(lSize - m_lLastX) > 2) Then
                                Resize
                                m_lLastX = lSize
                                RePaint
                                m_cTreeView.Refresh (m_eTvControlAlignment = etvRightAlign)
                                m_cTreeView.RePaint
                            End If
                        End If
                    End If
                Else
                    lSize = tPnt.y
                    If (lSize > 0) Then
                        GetClientRect m_lParentHwnd, tRect
                        If (lSize < (tRect.Bottom - 10)) Then
                            m_bTreeViewSizing = True
                            If (m_eTvControlAlignment = etvTopAlign) Then
                                m_cTreeView.Height = lSize
                            Else
                                m_cTreeView.Height = tRect.Bottom - lSize
                            End If
                            If (Abs(lSize - m_lLastX) > 2) Then
                                Resize
                                m_lLastX = lSize
                                RePaint
                                m_cTreeView.Refresh (m_eTvControlAlignment = etvBottomAlign)
                                m_cTreeView.RePaint
                            End If
                        End If
                    End If
                End If
            End If
            If Not (m_cSkinScrollBars Is Nothing) Then
                m_cSkinScrollBars.ByPassHitTest = False
            End If
        End If
    End If

End Sub

Private Function RowChecked(ByVal lRow As Long) As Boolean
'/* determine check state
    
    If ArrayCheck(m_cGridItem) Then
        RowChecked = m_cGridItem(lRow).Checked
    End If

End Function

Public Sub RowEnsureVisible(ByVal lRow As Long)
'*/ move to row index

    If Not (m_lHGHwnd = 0) Then
        If Not (lRow < 0) Then
            If Not (lRow > Count - 1) Then
                SendMessageLongA m_lHGHwnd, LVM_ENSUREVISIBLE, lRow, 0&
            End If
        End If
    End If
    m_cSkinScrollBars.Refresh

End Sub

Private Function RowHeightChange()
'/* force row height change

Dim tRect   As RECT
Dim tWPos   As WINDOWPOS

    GetWindowRect m_lHGHwnd, tRect
    With tRect
        OffsetRect tRect, -.left, -.top
    End With
    With tWPos
        .hwnd = m_lHGHwnd
        .hWndInsertAfter = 0&
        .x = 0&
        .y = 0&
        .flags = SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOOWNERZORDER Or SWP_NOZORDER
        .cx = tRect.Right
        .cy = tRect.Bottom
        SendMessageA m_lHGHwnd, WM_WINDOWPOSCHANGED, 0&, tWPos
    End With

End Function

Public Sub RowHideCheckBox(ByVal lRow As Long, _
                           ByVal bHide As Boolean)

'/* hide row checkbox

    If Not (m_lHGHwnd = 0) Then
        If Not m_bVirtualMode Then
            If (lRow < m_lRowCount) Then
                m_cGridItem(lRow).HideCheckBox = bHide
            End If
        End If
    End If

End Sub

Public Sub RowNoFocus(ByVal lRow As Long, _
                      ByVal bNoFocus As Boolean)

'/* no focus effect on row

    If Not (m_lHGHwnd = 0) Then
        If Not m_bVirtualMode Then
            If (lRow < m_lRowCount) Then
                m_cGridItem(lRow).RowNoFocus = bNoFocus
            End If
        End If
    End If

End Sub

Public Sub RowNoEdit(ByVal lRow As Long, _
                     ByVal bNoEdit As Boolean)

'/* no focus effect on row

    If Not (m_lHGHwnd = 0) Then
        If Not m_bVirtualMode Then
            If (lRow < m_lRowCount) Then
                m_cGridItem(lRow).RowNoEdit = bNoEdit
            End If
        End If
    End If

End Sub

Private Function RowHitTest() As Long
'/* row hit test

Dim lIndex  As Long
Dim tLVHT   As LVHITTESTINFO
Dim tPoint  As POINTAPI

    '/* get target item index
    GetCursorPos tPoint
    ScreenToClient m_lHGHwnd, tPoint
    RowHitTest = -1
    LSet tLVHT.pt = tPoint
    SendMessageA m_lHGHwnd, LVM_HITTEST, 0&, tLVHT
    If (tLVHT.iItem <= 0) Then
        If (tLVHT.flags And LVHT_NOWHERE) = LVHT_NOWHERE Then
            lIndex = -1
        Else
            lIndex = tLVHT.iItem
        End If
    Else
        lIndex = tLVHT.iItem
    End If
    If m_bUseSpannedRows Then
        lIndex = RowSpanMapVirtual(lIndex)
    End If
    RowHitTest = lIndex

End Function

Public Function CellIsVisible(ByVal lRow As Long, _
                              ByVal lCell As Long) As Boolean
'/* return row in client area

Dim tRClt   As RECT
Dim tRRow   As RECT
Dim tPnt    As POINTAPI

    CellCalcRect lRow, lCell, tRRow
    CopyMemory tPnt, tRRow, Len(tPnt)
    GetClientRect m_lHGHwnd, tRClt
    If Not (PtInRect(tRClt, tPnt.x, tPnt.y) = 0) Then
        CellIsVisible = True
    End If

End Function

Public Sub RowRedraw(ByVal lRow As Long)
'/* redraw row

    If Not (m_lHGHwnd = 0) Then
        If (lRow < m_lRowCount) Then
            SendMessageLongA m_lHGHwnd, LVM_REDRAWITEMS, lRow, lRow
        End If
    End If

End Sub

Public Function RowRemove(ByVal lRow As Long) As Boolean
'*/ remove an item from the list

On Error GoTo Handler

    If Not (m_lHGHwnd = 0) Then
        If Not m_bVirtualMode Then
            If (lRow < RowCount) Then
                If ArrayCheck(m_cGridItem) Then
                    '/* remove item
                    Set m_cGridItem(lRow) = Nothing
                    '/* reset array
                    ResizeArray m_cGridItem, lRow
                    '/* test for spanned
                    If m_bUseSpannedRows Then
                        RowArrayResize
                        m_lRowCount = m_lRowCount - 1
                        If (RowCount = 0) Then
                            SetRowCount 0
                            m_bHasInitialized = False
                        End If
                    Else
                        m_lRowCount = m_lRowCount - 1
                        SetRowCount m_lRowCount
                        If (m_lRowCount = 0) Then
                            m_bHasInitialized = False
                        End If
                    End If
                Else
                    SetRowCount 0
                    m_bHasInitialized = False
                End If
            Else
                m_bHasInitialized = False
            End If
        End If
    End If
    '/* success
    RowRemove = True
    RaiseEvent eVHItemDeleted(lRow)

Handler:
    On Error GoTo 0
    
End Function

Public Function RowsPerPage() As Long
'/* get first row item index

    If Not (m_lHGHwnd = 0) Then
        RowsPerPage = SendMessageLongA(m_lHGHwnd, LVM_GETCOUNTPERPAGE, 0&, 0&)
    End If

End Function

Public Function RowTopIndex() As Long
'/* get first row item index

    If Not (m_lHGHwnd = 0) Then
        RowTopIndex = SendMessageLongA(m_lHGHwnd, LVM_GETTOPINDEX, 0&, 0&)
    End If

End Function

Private Sub SetBorderStyle(ByVal lHwnd As Long, _
                           ByVal eStyle As EBSBorderStyle)
'/* change border style

    Select Case eStyle
    Case ebsNone
        WindowStyle lHwnd, GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME
        WindowStyle lHwnd, GWL_EXSTYLE, 0, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE
    Case ebsThin
        WindowStyle lHwnd, GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME
        WindowStyle lHwnd, GWL_EXSTYLE, WS_EX_STATICEDGE, WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE
    Case ebsThick
        WindowStyle lHwnd, GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME
        WindowStyle lHwnd, GWL_EXSTYLE, WS_EX_CLIENTEDGE, WS_EX_STATICEDGE Or WS_EX_WINDOWEDGE
    End Select

End Sub

Private Sub SetExtendedStyle(ByVal lStyle As Long, _
                             ByVal lStyleNot As Long)

'*/ change list extended style params

Dim lNewStyle   As Long

    If Not (m_lHGHwnd = 0) Then
        lNewStyle = SendMessageLongA(m_lHGHwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
        lNewStyle = lNewStyle And Not lStyleNot
        lNewStyle = lNewStyle Or lStyle
        SendMessageLongA m_lHGHwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, lNewStyle
    End If

End Sub

Public Sub SetRowCount(ByVal lRows As Long)
'/* dimension list to item count

    If Not (m_lHGHwnd = 0) Then
        SendMessageA m_lHGHwnd, LVM_SETITEMCOUNT, lRows, LVSICF_NOINVALIDATEALL
    End If
    If m_bVirtualMode Then
        m_lRowCount = lRows
    End If
    
End Sub

Private Sub SetStyle(ByVal lStyle As Long, _
                     ByVal lStyleNot As Long)

'*/ change list style params

Dim lNewStyle   As Long

    If Not (m_lHGHwnd = 0) Then
        If m_bIsNt Then
            lNewStyle = GetWindowLongW(m_lHGHwnd, GWL_STYLE)
        Else
            lNewStyle = GetWindowLongA(m_lHGHwnd, GWL_STYLE)
        End If
        lNewStyle = lNewStyle And Not lStyleNot
        lNewStyle = lNewStyle Or lStyle
        If m_bIsNt Then
            SetWindowLongW m_lHGHwnd, GWL_STYLE, lNewStyle
        Else
            SetWindowLongA m_lHGHwnd, GWL_STYLE, lNewStyle
        End If
        SetWindowPos m_lHGHwnd, 0&, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
    End If

End Sub

Private Sub ToolTipTrack()
'/* track tooltip state

Dim tRect   As RECT
Dim tPnt    As POINTAPI

    '/* stationary cursor
    GetCursorPos tPnt
    '/* in screen
    ScreenToClient m_lHGHwnd, tPnt
    GetClientRect m_lHGHwnd, tRect
    tRect.top = m_lHeaderHeight + 2
    If (PtInRect(tRect, tPnt.x, tPnt.y) = 0) Then
        StopTipTimer
        Exit Sub
    End If
    
    With tPnt
        If .x = m_lLastX Then
            If (.y = m_lLastY) Then
                m_lTipTimer = m_lTipTimer + 1
            Else
                m_lTipTimer = 0
            End If
        Else
            m_lTipTimer = 0
        End If
        m_lLastX = .x
        m_lLastY = .y
    End With
    
    '/* show tip
    If (m_lTipTimer > (m_lCellTipDelayTime * 10)) Then
        If Not m_bShowing Then
            m_cCellTips.DrawTip
            m_bShowing = True
        End If
    End If
    '/* destroy tip
    If (m_lTipTimer > ((m_lCellTipDelayTime + m_lCellTipVisibleTime) * 10)) Then
        If m_bShowing Then
            m_cCellTips.DestroyToolTip
            StopTipTimer
            m_bShowing = False
        End If
    End If

End Sub

Private Function TranslateColor(ByVal clr As OLE_COLOR, _
                                Optional ByVal hpal As Long = 0) As Long

'/* translate to ole color

    If OleTranslateColor(clr, hpal, TranslateColor) Then
        TranslateColor = -1
    End If

End Function

Public Function UnCheckAll() As Boolean
'*/ unmark all checkboxes

Dim lCt As Long

    If Not (m_lHGHwnd = 0) Then
        If ArrayCheck(m_cGridItem) Then
            For lCt = LBound(m_cGridItem) To UBound(m_cGridItem)
                m_cGridItem(lCt).Checked = False
            Next lCt
            GridRefresh False
        End If
    End If

End Function

Public Function VisibleItemCount() As Long
'/* items in view

    If Not (m_lHGHwnd = 0) Then
        VisibleItemCount = SendMessageLongA(m_lHGHwnd, LVM_GETCOUNTPERPAGE, 0&, 0&)
    End If

End Function

Private Sub WindowStyle(ByVal lHwnd As Long, _
                        ByVal lType As Long, _
                        ByVal lStyle As Long, _
                        ByVal lStyleNot As Long)

'/* set window style bits

Dim lNewStyle As Long

    If m_bIsNt Then
        lNewStyle = GetWindowLongW(lHwnd, lType)
    Else
        lNewStyle = GetWindowLongA(lHwnd, lType)
    End If
    lNewStyle = (lNewStyle And Not lStyleNot) Or lStyle
    If m_bIsNt Then
        SetWindowLongW lHwnd, lType, lNewStyle
    Else
        SetWindowLongA lHwnd, lType, lNewStyle
    End If
    SetWindowPos lHwnd, 0&, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED

End Sub


'**********************************************************************
'*                              CELLS
'**********************************************************************

Public Sub AddCell(ByVal lRow As Long, _
                   ByVal lColumn As Long, _
                   Optional ByVal sText As String, _
                   Optional ByVal lAlign As ECTTextAlignFlags, _
                   Optional ByVal lIconIdx As Long = -1, _
                   Optional ByVal lBackColor As Long = &HF8F8F8, _
                   Optional ByVal lForeColor As Long = -1, _
                   Optional ByVal oFont As StdFont, _
                   Optional ByVal lIndent As Long = 0, _
                   Optional ByVal lSpanRowDepth As Long = -1)

'/* grid class interface

On Error GoTo Handler

    If (lRow < 0) Then Exit Sub
    If m_bVirtualMode Then Exit Sub
    
    If (lRow = 0) Then
        '/* set minimum row height
        CalcMinRowHeight
    End If
    
    '/* resize griditems array
    If (lRow > UBound(m_cGridItem)) Then
        GridArrayResize lRow
    End If
    If (lColumn > m_cGridItem(lRow).Count) Then
        m_cGridItem(lRow).ResizeArray lColumn
    End If
    
    '/* horizontal spanning state
    If (lColumn = 0) Then
        m_cGridItem(lRow).SpanRowDepth = lSpanRowDepth
        If (lSpanRowDepth > 1) Then
            m_bUseSpannedRows = True
        End If
    End If
    
    '/* grid item properties pass though
    With m_cGridItem(lRow)
        .Text(lColumn) = sText
        .Align(lColumn) = lAlign
        .Icon(lColumn) = lIconIdx
        .BackColor(lColumn) = lBackColor
        .ForeColor(lColumn) = lForeColor
        .Indent(lColumn) = lIndent
    End With
    
    '/* lock and size first column
    If (lRow = 0) Then
        If (lColumn = 0) Then
            If m_bLockFirstColumn Then
                FirstColumnWidth lRow
            End If
        End If
    End If
    
    '/* cell font property
    If Not (oFont Is Nothing) Then
        m_cGridItem(lRow).FontHnd(lColumn) = CellAddFont(oFont)
        Set oFont = Nothing
    End If
    
    '/* add a single item
    If m_bDraw Then
        If m_bUseSpannedRows Then
            RowArrayResize
        Else
            If (lColumn = 0) Then
                SetRowCount UBound(m_cGridItem) + 1
            End If
        End If
        m_bHasInitialized = True
    Else
        '/* start drawing after last item loaded
        If (lRow = (m_lRowCount - 1)) Then
            If (lColumn = (m_lColumnCount - 1)) Then
                If m_bFastLoad Then
                    '/* engage vertical span tracking
                    RowArrayResize
                    '/* enable draing
                    m_bDraw = True
                    m_bFastLoad = False
                    '/* refresh list
                    GridRefresh False
                End If
                m_bHasInitialized = True
            End If
        End If
    End If

On Error GoTo 0
Exit Sub

Handler:
    On Error GoTo 0
    RaiseEvent eVHErrCond("AddCell", Err.Number)

End Sub

Private Sub CalcRowOffsets()

Dim lCt     As Long
Dim lDepth  As Long

    ReDim m_lRowDepth(0 To UBound(m_cGridItem))
    For lCt = 0 To UBound(m_cGridItem)
        If (m_cGridItem(lCt).SpanRowDepth > 1) Then
            lDepth = lDepth + (m_cGridItem(lCt).SpanRowDepth * m_lRowHeight)
        Else
            lDepth = lDepth + m_lRowHeight
        End If
        m_lRowDepth(lCt) = lDepth
    Next lCt

End Sub

Public Sub AddCellHeader(ByVal lRow As Long, _
                         ByVal lColumn As Long, _
                         ByVal sHeaderText As String, _
                         ByVal lHeaderForeColor As Long, _
                         Optional ByVal lHeaderFocusForeColor As Long, _
                         Optional ByVal oHeaderFont As StdFont, _
                         Optional ByVal lHeaderTextAlign As ECTTextAlignFlags, _
                         Optional ByVal bUseSpannedCell As Boolean = True, _
                         Optional ByVal lIndent As Long = 0)

Dim lUb As Long

On Error GoTo Handler

    If (lRow < m_lRowCount) Then
        If (lColumn < m_lColumnCount) Then
            '/* get subcell class count
            If Not ArrayCheck(m_cCellHeader) Then
                lUb = 0
            Else
                lUb = UBound(m_cCellHeader) + 1
            End If
            '/* create the instance
            ReDim Preserve m_cCellHeader(0 To lUb)
            Set m_cCellHeader(lUb) = New clsCellHeader
            With m_cCellHeader(lUb)
                .RowIndex = lRow
                .CellIndex = lColumn
                .AlignFlag = lHeaderTextAlign
                .ForeColor = lHeaderForeColor
                .FocusForeColor = lHeaderFocusForeColor
                .UseSpannedCell = bUseSpannedCell
                .Text = sHeaderText
                .Indent = lIndent
                If Not (oHeaderFont Is Nothing) Then
                    .FontHandle = CellAddFont(oHeaderFont)
                    Set oHeaderFont = Nothing
                Else
                    .FontHandle = -1
                End If
            End With
            m_cGridItem(lRow).CellHeader(lColumn) = lUb
        End If
    End If

On Error GoTo 0
Exit Sub

Handler:
    On Error GoTo 0
    RaiseEvent eVHErrCond("AddCellHeader", Err.Number)

End Sub


'> SubCell Edit Controls
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Public Sub SubCellAddControl(ByVal lRow As Long, _
                             ByVal lCell As Long, _
                             ByVal lWidth As Long, _
                             ByVal lHeight As Long, _
                             ByVal lCntlHwnd As Long, _
                             ByVal eFramePosition As EVSFramePosition, _
                             Optional ByVal eFrameConnector As EVSFrameConnector = 0, _
                             Optional ByVal lLeft As Long = 0, _
                             Optional ByVal lTop As Long = 0, _
                             Optional ByVal bSpanUseVirtualRow As Boolean = False, _
                             Optional ByVal bSpanUseVirtualCell As Boolean = False)

'/* add an edit cntl to subcell

Dim lUb As Long

On Error GoTo Handler

    If m_bVirtualMode Then Exit Sub
    If (m_lHGHwnd = 0) Then Exit Sub
    
    If (lRow < m_lRowCount) Then
        If (lCell < m_lColumnCount) Then
            If Not (lCntlHwnd = 0) Then
                '/* get subcell class count
                If Not ArrayCheck(m_cSubCell) Then
                    lUb = 0
                Else
                    lUb = UBound(m_cSubCell) + 1
                End If
                '/* create the instance
                ReDim Preserve m_cSubCell(0 To lUb)
                Set m_cSubCell(lUb) = New clsSubCell
                '/* create control dc
                ReDim Preserve m_cControlDc(0 To lUb)
                Set m_cControlDc(lUb) = New clsStoreDc
                With m_cSubCell(lUb)
                    .RowIndex = lRow
                    .CellIndex = lCell
                    .Width = lWidth
                    .Height = lHeight
                    .CntlHwnd = lCntlHwnd
                    .FramePosition = eFramePosition
                    .FrameConnector = eFrameConnector
                    .left = lLeft
                    .top = lTop
                    .SpanUseVirtualRow = bSpanUseVirtualRow
                    .SpanUseVirtualCell = bSpanUseVirtualCell
                End With
                '/* flag row
                m_cGridItem(lRow).HasSubCells = True
                m_cGridItem(lRow).SubCellInstance = lUb
                '/* subcell flag on
                m_bHasSubCells = True
            End If
        End If
    End If

On Error GoTo 0
Exit Sub

Handler:
    On Error GoTo 0
    RaiseEvent eVHErrCond("AddCellHeader", Err.Number)

End Sub

Private Sub SubCellsHideAll()

Dim lUb     As Long
Dim lCt     As Long
Dim lHwnd   As Long

    lUb = UBound(m_cSubCell)
    For lCt = 0 To lUb
        '/* get cntl handle
        lHwnd = m_cSubCell(lCt).CntlHwnd
        If Not (lHwnd = 0) Then
            SubCellShowControl lHwnd, False
        End If
    Next lCt
    
End Sub

Private Sub SubCellScrollCntl(ByVal lHdc As Long)
'/* cntl state hub

Dim lUb     As Long
Dim lCt     As Long
Dim lRow    As Long
Dim lCell   As Long
Dim lHwnd   As Long
Dim tRCell  As RECT

    lUb = UBound(m_cSubCell)
    For lCt = 0 To lUb
        '/* get cntl handle
        lHwnd = m_cSubCell(lCt).CntlHwnd
        If Not (lHwnd = 0) Then
            '/* if spanned get virtual address
            With m_cSubCell(lCt)
                '/* use row coords
                If m_bUseSpannedRows Then
                    '/* draw to actual row
                    If Not .SpanUseVirtualRow Then
                        lRow = RowSpanMapVirtual(.RowIndex)
                    Else
                        lRow = .RowIndex
                    End If
                Else
                    lRow = .RowIndex
                End If
                If .SpanUseVirtualCell Then
                    If CellIsSpanned(lRow, .CellIndex) Then
                        lCell = CellSpanFirstCell(lRow)
                    Else
                        lCell = .CellIndex
                    End If
                Else
                    lCell = .CellIndex
                End If
                '/* row not in view
                If Not (CellIsVisible(lRow, lCell)) Then
                    '/* hide cntl
                    If SubCellCntlVisible(lHwnd) Then
                        SubCellShowControl lHwnd, False
                    End If
                Else
                    '/* get the cell rect
                    CellCalcRect lRow, lCell, tRCell
                    '/* calculate cntl position offsets
                    SubCellRect .left, .top, .Width, .Height, .FramePosition, tRCell
                    '/* paint conditionals
                    If SubCellInTransition Then
                        SubCellShowControl lHwnd, False
                        SubCellPaintControl lCt, lHdc, tRCell
                    ElseIf (tRCell.top < (m_lHeaderOffset)) Then
                        SubCellShowControl lHwnd, False
                    ElseIf (tRCell.left < 0) Then
                        SubCellShowControl lHwnd, False
                    Else
                        '/* position cntl
                        SubCellMoveControl lHwnd, tRCell
                        SubCellSizeControl lHwnd, tRCell
                        '/* get backcolor
                        SubCellCntlBackColor lHwnd
                        '/* show
                        SubCellShowControl lHwnd, True
                        '/* store cntl bmp
                        SubCellStoreBitmap lCt, tRCell
                    End If
                End If
            End With
        End If
    Next lCt

End Sub

Private Sub SubCellSortRows()
'/* reassign sorted subcell class pointers

Dim lUb As Long
Dim lCt As Long
Dim lCs As Long

    lUb = UBound(m_cGridItem)
    For lCt = 0 To lUb
        With m_cGridItem(lCt)
            If .HasSubCells Then
                For lCs = 0 To .SubCellCount
                    m_cSubCell(lCs).RowIndex = lCt
                Next lCs
            End If
        End With
    Next lCt

End Sub

Private Sub SubCellStoreBitmap(ByVal lCIdx As Long, _
                               ByRef tRect As RECT)

'/* store cntl image for transition painting

Dim lHwnd   As Long
Dim lHdc    As Long

    lHwnd = m_cSubCell(lCIdx).CntlHwnd
    If Not (lHwnd = 0) Then
        SubCellCntlBackColor lHwnd
        '/* cntl visible
        If SubCellCntlVisible(lHwnd) Then
            '/* refresh cntl
            UpdateWindow lHwnd
            '/* cntl dc
            lHdc = GetDC(lHwnd)
            '/* create dc
            Set m_cControlDc(lCIdx) = Nothing
            Set m_cControlDc(lCIdx) = New clsStoreDc
            With tRect
                m_cControlDc(lCIdx).Width = (.Right - .left)
                m_cControlDc(lCIdx).Height = (.Bottom - .top)
                '/* copy dc
                m_cRender.Blit m_cControlDc(lCIdx).hdc, 0, 0, (.Right - .left), (.Bottom - .top), lHdc, 0, 0, SRCCOPY
            End With
            '/* release
            ReleaseDC lHwnd, lHdc
        End If
    End If

End Sub

Private Sub SubCellPaintControl(ByVal lCIdx As Long, _
                                ByVal lHdc As Long, _
                                ByRef tRCntl As RECT)

'/* substitute cntl image while grid transitioning

Dim lCtDc   As Long
Dim lHwnd   As Long
Dim lRgn    As Long

    
    lCtDc = m_cControlDc(lCIdx).hdc
    lHwnd = m_cSubCell(lCIdx).CntlHwnd
    If Not (lCtDc = 0) Then
        '/* paint cntl image
        With tRCntl
            lRgn = CreateRectRgn(.left, .top, .Right, .Bottom)
            SelectClipRgn lHdc, lRgn
            m_cRender.Blit lHdc, .left, .top, (.Right - .left) + 10, (.Bottom - .top), lCtDc, 0, 0, SRCCOPY
            SelectClipRgn lHdc, 0&
            DeleteObject lRgn
        End With
    '/* position cntl
    Else
        SubCellMoveControl lHwnd, tRCntl
        SubCellShowControl lHwnd, True
    End If

End Sub

Private Function SubCellCntlVisible(ByVal lHwnd As Long) As Boolean
'/* return control visibility
    SubCellCntlVisible = CBool(IsWindowVisible(lHwnd))
End Function

Private Sub SubCellShowControl(ByVal lHwnd As Long, _
                               ByVal bVisible As Boolean)

'/* toggle edit cntl visibility

    If bVisible Then
        If Not (SubCellCntlVisible(lHwnd)) Then
            ShowWindow lHwnd, SW_NORMAL
        End If
    Else
        If SubCellCntlVisible(lHwnd) Then
            ShowWindow lHwnd, SW_HIDE
        End If
    End If

End Sub

Private Sub SubCellMoveControl(ByVal lHwnd As Long, _
                               ByRef tRCell As RECT)
'/* move edit cntl

    With tRCell
        SetWindowPos lHwnd, 0&, .left, .top, 0&, 0&, SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOSIZE
    End With

End Sub

Private Sub SubCellSizeControl(ByVal lHwnd As Long, _
                               ByRef tRCell As RECT)
'/* size edit cntl

    With tRCell
        SetWindowPos lHwnd, 0&, 0&, 0&, (.Right - .left), (.Bottom - .top), SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOMOVE
    End With

End Sub

Private Sub SubCellCntlBackColor(ByVal lHwnd As Long)

Dim lHdc    As Long
Dim lStyle  As Long

    If Not (lHwnd = 0) Then
        '/* test for od button style
        lStyle = GetWindowLongA(lHwnd, GWL_STYLE)
        '/* send clr change to trigger repaint sub
        If (lStyle And &HB) = &HB Then
            lHdc = GetDC(lHwnd)
            SendMessageLongA m_lHGHwnd, WM_CTLCOLORBTN, lHdc, lHwnd
            ReleaseDC lHwnd, lHdc
            UpdateWindow lHwnd
        ElseIf (lStyle And &HD) = &HD Then
            lHdc = GetDC(lHwnd)
            SendMessageLongA m_lHGHwnd, WM_CTLCOLORSTATIC, lHdc, lHwnd
            ReleaseDC lHwnd, lHdc
            UpdateWindow lHwnd
        End If
    End If

End Sub

Private Function SubCellInTransition() As Boolean
'/* grid view state

    Select Case True
    Case m_bColumnDragging, m_bColumnSizingVertical, m_bColumnSizingHorizontal, m_bEditorLoaded, m_bItemActive, m_bRowDragging
        SubCellInTransition = True
    Case Else
        If (m_bVisible = False) Then
            SubCellInTransition = True
        ElseIf (m_bEnabled = False) Then
            SubCellInTransition = True
        ElseIf ScrollState Then
            SubCellInTransition = True
        Else
            SubCellInTransition = False
        End If
    End Select

End Function

Private Sub SubCellRect(ByVal lLeft As Long, _
                        ByVal lTop As Long, _
                        ByVal lWidth As Long, _
                        ByVal lHeight As Long, _
                        ByVal ePosition As EVSFramePosition, _
                        ByRef tRCell As RECT)

'/* calculate relative cntl position within cell

Dim tRClt As RECT

    With tRCell
        If ((lWidth + lLeft + 4) > (.Right - .left)) Then
            lWidth = (.Right - .left) - 4
            lLeft = 2
        End If
        If ((lHeight + lTop + 4) > (.Bottom - .top)) Then
            lHeight = (.Bottom - .top) - 4
            lTop = 2
        End If
        Select Case ePosition
        Case evsUserDefine
            .left = .left + lLeft
            .top = .top + lTop
            .Right = .left + lWidth
            .Bottom = .top + lHeight
        Case evsBottomLeft
            .left = .left + 2
            .Right = .left + lWidth
            .top = .Bottom - (lHeight + 2)
            .Bottom = .top + lHeight
        Case evsBottomCenter
            .left = .left + (((.Right - .left) - lWidth) / 2)
            .Right = .left + lWidth
            .top = .Bottom - (lHeight + 2)
            .Bottom = .top + lHeight
        Case evsBottomRight
            .left = .Right - (lWidth + lLeft)
            .Right = .left + lWidth
            .top = .Bottom - (lHeight + 2)
            .Bottom = .top + (lHeight)
        Case evsCenterLeft
            .left = .left + 2
            .Right = .left + lWidth
            .top = .top + (((.Bottom - .top) - lHeight) / 2)
            .Bottom = .top + lHeight
        Case evsCenterCell
            .left = .left + (((.Right - .left) - lWidth) / 2)
            .Right = (.left + lWidth)
            .top = .top + (((.Bottom - .top) - lHeight) / 2)
            .Bottom = (.top + lHeight)
        Case evsCenterRight
            .left = .Right - (lWidth + 2)
            .Right = (.left + lWidth)
            .top = .top + (((.Bottom - .top) - lHeight) / 2)
            .Bottom = (.top + lHeight)
        Case evsTopLeft
            .left = .left + 2
            .Right = .left + lWidth
            .top = .top + 2
            .Bottom = .top + lHeight
        Case evsTopCenter
            .left = .left + (((.Right - .left) - lWidth) / 2)
            .Right = .left + lWidth
            .top = .top + 2
            .Bottom = .top + lHeight
        Case evsTopRight
            .left = .Right - (lWidth + 2)
            .Right = .left + lWidth
            .top = .top + 2
            .Bottom = .top + lHeight
        End Select
        GetClientRect m_lHGHwnd, tRClt
        If (.Right > (tRClt.Right - 2)) Then
            .Right = tRClt.Right - 2
        End If
    End With

End Sub

Private Function ScrollState() As Boolean
'/* grid scroll state
    If Not (m_cSkinScrollBars Is Nothing) Then
        ScrollState = Not (m_cSkinScrollBars.ScrollDirection = efsNone)
    End If
End Function

'> Cell Spanning
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Function CellSpanFirstCell(ByVal lRow As Long) As Long

    With m_cGridItem(lRow)
        If Not (.SpanFirstCell = -1) Then
            CellSpanFirstCell = .SpanFirstCell
        End If
    End With

End Function

Private Function CellSpanHasCandidate(ByVal lRow As Long) As Boolean

Dim lCt As Long

    With m_cGridItem(lRow)
        For lCt = 0 To .Count
            If Not (.SpanFirstCell = -1) Then
                CellSpanHasCandidate = True
                Exit For
            End If
        Next lCt
    End With

End Function

Public Sub CellSpanHorizontal(ByVal lRow As Long, _
                              ByVal lFirstCell As Long, _
                              ByVal lLastCell As Long)
'/* add a spanned cell

    If Not m_bVirtualMode Then
        m_cGridItem(lRow).CellSpanHorizontal lFirstCell, lLastCell
    End If

End Sub

Public Sub CellSpanVertical(ByVal lRow As Long, _
                            ByVal lSpanDepth As Long)
'/* add a spanned row

    If Not m_bVirtualMode Then
        m_cGridItem(lRow).SpanRowDepth = lSpanDepth
        RowArrayResize
    End If

End Sub

Private Sub CalcCheckBoxRect(ByRef tRect As RECT)
'/* calculate checkbox position

Dim lYOffset As Long

    With tRect
        lYOffset = ((.Bottom - .top) - 13) / 2
        .top = .top + lYOffset + 1
        .Bottom = .top + 13
        .left = .left + 4
        .Right = .left + 13
    End With

End Sub

'> Cell Rendering
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Function CalcMinRowHeight() As Long
'/* calculate minimum row height

    If (m_lImlRowHndl > 0) Then
        m_lRowMinHgt = m_lRowIconX + 4
    Else
        m_lRowMinHgt = 24
    End If
    If (m_lRowHeight < m_lRowMinHgt) Then
        m_lRowHeight = m_lRowMinHgt
        RowHeightChange
    End If

End Function

Private Function CellAddFont(ByVal oFont As StdFont) As Long
'/* add a cell font to font array

Dim lCt As Long
Dim lUb As Long

On Error Resume Next

    If (oFont Is Nothing) Then Exit Function
    '/* test for an existing match and add font handle
    If ArrayCheck(m_oCellFont) Then
        lUb = UBound(m_oCellFont)
        If IsError(m_oCellFont(lUb).Name) Then
            lUb = 0
            ReDim m_oCellFont(0)
            ReDim m_hFontHnd(0)
        Else
            For lCt = 0 To lUb
                If (oFont.Name = m_oCellFont(lCt).Name) Then
                    If (oFont.Bold = m_oCellFont(lCt).Bold) Then
                        If (oFont.Italic = m_oCellFont(lCt).Italic) Then
                            If (oFont.Underline = m_oCellFont(lCt).Underline) Then
                                If (oFont.Size = m_oCellFont(lCt).Size) Then
                                    If (oFont.Strikethrough = m_oCellFont(lCt).Strikethrough) Then
                                        CellAddFont = m_hFontHnd(lCt)
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next lCt
            lUb = lUb + 1
        End If
    Else
        lUb = 0
        ReDim m_oCellFont(0)
        ReDim m_hFontHnd(0)
    End If

    ReDim Preserve m_hFontHnd(0 To lUb) As Long
    ReDim Preserve m_oCellFont(0 To lUb) As StdFont
    Set m_oCellFont(lUb) = New StdFont
    Set m_oCellFont(lUb) = oFont
    m_hFontHnd(lUb) = FontHandle(m_oCellFont(lUb))
    If Not (m_hFontHnd(lUb) = 0) Then
        CellAddFont = m_hFontHnd(lUb)
    End If

On Error GoTo 0

End Function

Private Sub CellBackColor(ByVal lHdc As Long, _
                          ByVal lColor As Long, _
                          ByRef tRect As RECT)
'/* paint cell backcolor

Dim lhBrush As Long

    If (lColor = -1) Then
        lhBrush = GetSysColorBrush(UserControl.BackColor And &H1F&)
    ElseIf (lColor = 0) Then
        lhBrush = GetSysColorBrush(UserControl.BackColor And &H1F&)
    ElseIf m_bEnabled Then
        lhBrush = CreateSolidBrush(TranslateColor(lColor))
    Else
        lhBrush = CreateSolidBrush(TranslateColor(m_lDisabledBackColor))
    End If
    FillRect lHdc, tRect, lhBrush
    DeleteObject lhBrush

End Sub

Private Function ColumnsAreShuffled() As Boolean

Dim lCt As Long

    For lCt = 0 To (m_lColumnCount - 1)
        If Not (lCt = ColumnIndex(lCt)) Then
            ColumnsAreShuffled = True
            Exit For
        End If
    Next lCt

End Function

Private Function CellCalcRect(ByVal lRow As Long, _
                              ByVal lCell As Long, _
                              ByRef tRSub As RECT) As Long
'/* calculate cell dimensions

Dim tRcpy   As RECT

    '/* calculate horizontal base
    If (lCell = -1) Then
        GetRowRect lRow, tRSub
    Else
        If m_bVirtualMode Then
            GetCellRect lRow, lCell, tRSub
        Else
            If CellIsSpanned(lRow, lCell) Then
                With m_cGridItem(lRow)
                    If (lCell = .SpanFirstCell) Then
                        GetCellRect lRow, ColumnAtIndex(lCell), tRSub
                        GetCellRect lRow, ColumnAtIndex(.SpanLastCell), tRcpy
                        tRSub.Right = tRcpy.Right
                    Else
                        GetCellRect lRow, lCell, tRSub
                    End If
                End With
            Else
                If ColumnsAreShuffled Then
                    If CellSpanHasCandidate(lRow) Then
                        GetCellRect lRow, ColumnIndex(lCell), tRSub
                    Else
                        GetCellRect lRow, lCell, tRSub
                    End If
                Else
                    GetCellRect lRow, lCell, tRSub
                End If
            End If
        End If
    End If

End Function

Private Sub CellCalcHeaderSize(ByVal lRow As Long, _
                               ByVal lCell As Long, _
                               ByVal lHdc As Long, _
                               ByVal sText As String, _
                               ByRef tRect As RECT)

Dim lOldFont    As Long
Dim lFlags      As Long
Dim lDist       As Long
Dim lClHdr      As Long
Dim tRTxt       As RECT

    If Not m_bVirtualMode Then
        CopyRect tRTxt, tRect
        lClHdr = m_cGridItem(lRow).CellHeader(lCell)
        If (lClHdr > -1) Then
            '/* calc cell header
            If (LenB(m_cCellHeader(lClHdr).Text) > 0) Then
                With m_cCellHeader(lClHdr)
                    lFlags = .AlignFlag
                    sText = .Text
                    If Not (.FontHandle = 0) Then
                        lOldFont = SelectObject(lHdc, .FontHandle)
                    End If
                End With
                sText = sText & Chr(0)
                '/* get size
                If Not (tRect.left = 0) Then
                    tRTxt.left = tRect.left
                    InflateRect tRTxt, -2, 0&
                End If
                '/* get header rect
                If m_bIsNt Then
                    DrawTextW lHdc, StrPtr(sText), -1, tRTxt, DT_CALCRECT Or DT_SINGLELINE
                Else
                    DrawTextA lHdc, sText, -1, tRTxt, DT_CALCRECT Or DT_SINGLELINE
                End If
                If Not (lOldFont = 0) Then
                    SelectObject lHdc, lOldFont
                    lOldFont = 0
                End If
                lDist = (tRTxt.Bottom - tRTxt.top) + 2
                OffsetRect tRect, 0&, lDist
                lFlags = m_cGridItem(lRow).Align(lCell)
                If (lFlags And DT_VCENTER) = DT_VCENTER Then
                    If Not ((lFlags And DT_VCENTER) = DT_WORDBREAK) Then
                        OffsetRect tRect, 0&, -6
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Function CellColor(ByVal lRow As Long, _
                           ByVal lCell As Long) As Long

'/* apply row color patterns

    '/* test valid
    If Not ArrayCheck(m_lCellColor) Then
        CellDecoration m_eCellDecoration, m_lCellColorBase, m_lCellColorOffset, m_bCellUseXP, m_lCellDepth
    Else
        If m_bVirtualMode Then
            If (lRow > (Count - 1)) Then
                CellDecoration m_eCellDecoration, m_lCellColorBase, m_lCellColorOffset, m_bCellUseXP, m_lCellDepth, True
            End If
        Else
            If (lRow < LBound(m_lCellColor)) Or (lRow > UBound(m_lCellColor)) Then
                CellDecoration m_eCellDecoration, m_lCellColorBase, m_lCellColorOffset, m_bCellUseXP, m_lCellDepth, True
            End If
        End If
    End If

    '/* get color to geometric match patterns
    Select Case m_eCellDecoration
    Case erdCellLine
        CellColor = m_lCellColor(lRow)
    
    Case erdCellChecker
        If (lCell Mod 2) Then
            If m_lCellColor(lRow) = m_lCellColorBase Then
                CellColor = m_lCellColorOffset
            Else
                CellColor = m_lCellColorBase
            End If
        Else
            CellColor = m_lCellColor(lRow)
        End If
    
    Case erdCellBiLinear
        If (lCell Mod 2) Then
            CellColor = m_lCellColorOffset
        Else
            CellColor = m_lCellColorBase
        End If
    
    Case erdCellSplit
        If (lCell Mod 3) Then
            CellColor = m_lCellColorOffset
        Else
            CellColor = m_lCellColorBase
        End If
    End Select

End Function

Public Function CellDecoration(ByVal eCellDecoration As ERDCellDecoration, _
                               ByVal lBaseClr As Long, _
                               ByVal lOffsetClr As Long, _
                               ByVal bXPColors As Boolean, _
                               Optional ByVal lCellDepth As Long, _
                               Optional ByVal bResize As Boolean) As Boolean

'/* build custom row color arrays

Dim bCs     As Boolean
Dim lCt     As Long
Dim lCr     As Long
Dim lUb     As Long

On Error GoTo Handler

    If m_bVirtualMode Then
        lUb = (Count - 1)
    Else
        lUb = (RowCount - 1)
    End If
    m_eCellDecoration = eCellDecoration
    m_lCellDepth = lCellDepth
    m_bCellUseXP = bXPColors
    
    If Not bResize Then
        If bXPColors Then
            m_lCellColorBase = m_cRender.XPShift(lBaseClr, 120)
            m_lCellColorOffset = m_cRender.XPShift(lOffsetClr, 120)
        Else
            m_lCellColorBase = lBaseClr
            m_lCellColorOffset = lOffsetClr
        End If
    End If
    If (lUb = -1) Then
        m_bCellDecoration = True
        Exit Function
    End If
    
    ReDim m_lCellColor(lUb)
    Do
        If (lCr > lCellDepth) Then
            lCr = 0
            bCs = Not bCs
        End If
        If bCs Then
            m_lCellColor(lCt) = m_lCellColorBase
        Else
            m_lCellColor(lCt) = m_lCellColorOffset
        End If
        lCr = lCr + 1
        lCt = lCt + 1
    Loop Until (lCt > lUb)

    m_bCellDecoration = True

Handler:

End Function

Private Function CellDrawCheckBox(ByVal lRow As Long, _
                                  ByVal lHdc As Long, _
                                  ByRef tRect As RECT, _
                                  Optional bDisabled As Boolean) As Boolean

'/* render custom checkboxes

Dim lXOffset As Long
Dim lYOffset As Long
Dim lState   As Long
Dim lFlags   As Long
Dim lCkPtr   As Long

    lFlags = ILD_TRANSPARENT
    With tRect
        lCkPtr = lRow
        '/* image state
        If Not m_bEnabled Then
            lState = 2
        ElseIf bDisabled Then
            lState = 2
        ElseIf RowChecked(lCkPtr) Then
            lState = 1
        Else
            lState = 0
        End If
        '/* horizontal offset
        lYOffset = ((.Bottom - .top) - 13) / 2
        lXOffset = 3
        '/* draw checkbox
        ImageList_Draw m_lImlStateHndl, lState, lHdc, (.left + lXOffset), (.top + lYOffset), lFlags
    End With

End Function

Private Sub CellDrawIcon(ByVal lHdc As Long, _
                         ByVal lIcon As Long, _
                         ByVal lState As Long, _
                         ByRef tRect As RECT)

'/* draw the cell icon

Dim lY      As Long
Dim lX      As Long
Dim lRgn    As Long

    '/* draw image
    With tRect
        Select Case m_lIconPosition
        '/* center align
        Case 0
            .top = .top + (((.Bottom - .top) - m_lRowIconY) / 2)
            .Bottom = .top + m_lRowIconY
            .left = .left + 1
            .Right = .left + m_lRowIconX
        '/* top align
        Case 1
            .top = .top + 1
            .Bottom = .top + m_lRowIconY
            .left = .left + 1
            .Right = .left + m_lRowIconX
        '/* bottom align
        Case 2
            .top = (.Bottom - m_lRowIconY) + 1
            .Bottom = .top + m_lRowIconY
            .left = .left + 1
            .Right = .left + m_lRowIconX
        End Select
        
        If Not (m_lImlRowHndl = 0) Then
            '/* create clipping region
            lRgn = CreateRectRgn(.left, .top, .Right, .Bottom)
            SelectClipRgn lHdc, lRgn
            '/* draw icon
            If m_bEnabled Then
                If (lState = 3) Then
                    If Not m_bIconNoHilite Then
                        m_cCellIcons.DrawImage lHdc, lIcon, (.left + lX), (.top + lY), ildSelected
                    Else
                        m_cCellIcons.DrawImage lHdc, lIcon, (.left + lX), (.top + lY)
                    End If
                Else
                    m_cCellIcons.DrawImage lHdc, lIcon, (.left + lX), (.top + lY)
                End If
            Else
                m_cCellIcons.DrawImage lHdc, lIcon, (.left + lX), (.top + lY), ildDisabled
            End If
        End If
    End With
    
    '/* cleanup
    SelectClipRgn lHdc, 0&
    DeleteObject lRgn

End Sub

Private Function CalcTextBestFit(ByVal lHdc As Long, _
                                 ByRef tRect As RECT) As Long

Dim lHeight As Long
Dim lWidth  As Long
Dim tPnt    As POINTAPI

On Error GoTo Handler

    CharDimensions lHdc, tPnt
    If (tPnt.x = 0) Then
        Exit Function
    ElseIf (tPnt.y = 0) Then
        Exit Function
    End If
    With tRect
        lWidth = (.Right - .left) / tPnt.x
        lHeight = (.Bottom - .top) / tPnt.y
    End With
    CalcTextBestFit = lWidth * lHeight

Handler:
On Error GoTo 0

End Function

Private Sub CellDrawText(ByVal lRow As Long, _
                         ByVal lCell As Long, _
                         ByVal lHdc As Long, _
                         ByVal sText As String, _
                         ByRef tRect As RECT)
'/* draw cell text
Dim lTrim       As Long
Dim lOldFont    As Long
Dim lFlags      As Long
Dim lDist       As Long
Dim lClHdr      As Long
Dim tRTxt       As RECT
Dim tRHdr       As RECT

    If Not m_bVirtualMode Then
        lClHdr = m_cGridItem(lRow).CellHeader(lCell)
        If (lClHdr > -1) Then
            If (LenB(m_cCellHeader(lClHdr).Text) > 0) Then
                If (m_cGridItem(lRow).Icon(lCell) > -1) Then
                    CellCalcRect lRow, lCell, tRHdr
                End If
                CellDrawHeader lClHdr, lHdc, tRHdr.left, lRow, lCell, tRect
            End If
        End If
        With m_cGridItem(lRow)
            '/* extended formatting
            lFlags = .Align(lCell)
            '/* cell font
            If Not (.FontHnd(lCell) = 0) Then
                lOldFont = SelectObject(lHdc, .FontHnd(lCell))
            End If
        End With
    End If
    
    If (LenB(sText) = 0) Then
        GoTo Handler
    End If
    SetBkMode lHdc, BM_TRANSPARENT
    CopyRect tRTxt, tRect
    With tRTxt
        .Right = .Right - 2
        .Bottom = .Bottom - 4
    End With
    If m_bIsNt Then
        DrawTextW lHdc, StrPtr(sText), -1, tRTxt, lFlags Or DT_CALCRECT
    Else
        DrawTextA lHdc, sText, -1, tRTxt, lFlags Or DT_CALCRECT
    End If

    '/* formatting flags
    With tRTxt
        If (.Right > tRect.Right) Then
            .Right = tRect.Right
        End If
        If (.Bottom > tRect.Bottom) Then
            .Bottom = tRect.Bottom
        End If
        If (lFlags And DT_VCENTER) = DT_VCENTER Then
            lDist = (tRect.Bottom - .Bottom) / 2
            If (.Bottom > tRect.Bottom) Then
                lDist = lDist - (tRTxt.Bottom - (tRect.Bottom + 2))
            End If
            OffsetRect tRTxt, 0&, lDist
        ElseIf (lFlags And DT_BOTTOM) = DT_BOTTOM Then
            lDist = tRect.Bottom - .Bottom
            OffsetRect tRTxt, 0&, lDist
        End If
        If (lFlags And DT_RIGHT) = DT_RIGHT Then
            lDist = tRect.Right - .Right
            OffsetRect tRTxt, lDist, 0&
        ElseIf (lFlags And DT_CENTER) = DT_CENTER Then
            lDist = (tRect.Right - .Right) / 2
            OffsetRect tRTxt, lDist, 0&
        End If
        If (lFlags And DT_WORDBREAK = DT_WORDBREAK) Then
            On Error Resume Next
            lTrim = (CalcTextBestFit(lHdc, tRTxt))
            If (Len(sText) > lTrim) Then
                lTrim = InStr(lTrim, sText, " ")
                If Not (lTrim < 1) Then
                    sText = left$(sText, lTrim) & ".."
                End If
            End If
            On Error GoTo 0
        End If
    End With
    sText = sText & vbNullChar
    '/* draw text
    If m_bIsNt Then
        If m_bFontRightLeading Then
            lFlags = lFlags Or DT_RTLREADING
        End If
        DrawTextW lHdc, StrPtr(sText), -1, tRTxt, lFlags Or DT_WORD_ELLIPSIS 'Or DT_NOCLIP
    Else
        DrawTextA lHdc, sText, -1, tRTxt, lFlags Or DT_WORD_ELLIPSIS 'Or DT_NOCLIP
    End If

Handler:
    '/* cleanup
    If Not (lOldFont = 0) Then
        SelectObject lHdc, lOldFont
        lOldFont = 0
    End If

End Sub
           
Public Function CellHeaderText(ByVal lRow As Long, _
                               ByVal lCell As Long) As String

Dim lClHdr  As Long

    If Not m_bVirtualMode Then
        lClHdr = m_cGridItem(lRow).CellHeader(lCell)
        If (lClHdr > -1) Then
            CellHeaderText = m_cCellHeader(lClHdr).Text
        End If
    End If

End Function

Private Sub CellDrawHeader(ByVal lClHdr As Long, _
                           ByVal lHdc As Long, _
                           ByVal lLeft As Long, _
                           ByVal lRow As Long, _
                           ByVal lCell As Long, _
                           ByRef tRect As RECT)

Dim bFocus      As Boolean
Dim lFlags      As Long
Dim lOldFont    As Long
Dim lForeClr    As Long
Dim lFocusClr   As Long
Dim lOldColor   As Long
Dim lDist       As Long
Dim sText       As String
Dim tRHdr       As RECT

    With m_cCellHeader(lClHdr)
        lFlags = .AlignFlag
        lForeClr = .ForeColor
        lFocusClr = .FocusForeColor
        sText = .Text
        lLeft = lLeft + .Indent
        If Not (.FontHandle = 0) Then
            lOldFont = SelectObject(lHdc, .FontHandle)
        End If
        '/* get size
        If Not .UseSpannedCell Then
            GetCellRect lRow, lCell, tRHdr
        Else
            CopyRect tRHdr, tRect
        End If
    End With
    If (LenB(sText) = 0) Then
        Exit Sub
    Else
        sText = sText & Chr(0)
    End If
    
    If Not m_bGridFocused Then
        bFocus = False
    ElseIf m_bFocusTextOnly Then
        bFocus = False
    Else
        If m_bFullRowSelect Then
            bFocus = RowFocused(lRow)
        Else
            bFocus = CellFocused(lRow, lCell)
        End If
    End If
    '/* text color
    If Not m_bEnabled Then
        lForeClr = m_lDisabledForeColor
    ElseIf bFocus Then
        If Not (lFocusClr = -1) Then
            lOldColor = GetTextColor(lHdc)
            CellForeColor lHdc, lFocusClr
        Else
            If Not (lForeClr = -1) Then
                lOldColor = GetTextColor(lHdc)
                CellForeColor lHdc, lForeClr
            End If
        End If
    Else
        If Not (lForeClr = -1) Then
            lOldColor = GetTextColor(lHdc)
            CellForeColor lHdc, lForeClr
        End If
    End If

    If Not (lLeft = 0) Then
        tRHdr.left = lLeft
        InflateRect tRHdr, -2, 0&
    End If
    If m_bIsNt Then
        DrawTextW lHdc, StrPtr(sText), -1, tRHdr, DT_CALCRECT Or DT_SINGLELINE
    Else
        DrawTextA lHdc, sText, -1, tRHdr, DT_CALCRECT Or DT_SINGLELINE
    End If
    
    '/* formatting flags
    With tRHdr
        If (.Right > (tRect.Right - 2)) Then
            .Right = tRect.Right - 2
        End If
        If (lFlags And DT_RIGHT) = DT_RIGHT Then
            lDist = tRect.Right - .Right
            OffsetRect tRHdr, lDist, 0&
        ElseIf (lFlags And DT_CENTER) = DT_CENTER Then
            lDist = (tRect.Right - .Right) / 2
            OffsetRect tRHdr, lDist, 0&
        End If
    End With
    
    SetBkMode lHdc, BM_TRANSPARENT
    '/* draw text
    If m_bIsNt Then
        If m_bFontRightLeading Then
            lFlags = lFlags Or DT_RTLREADING
        End If
        DrawTextW lHdc, StrPtr(sText), -1, tRHdr, DT_SINGLELINE Or DT_WORD_ELLIPSIS
    Else
        DrawTextA lHdc, sText, -1, tRHdr, DT_SINGLELINE Or DT_WORD_ELLIPSIS
    End If
    '/* adjust rect
    tRect.top = tRHdr.Bottom + 2
    '/* restore text color
    If Not (lForeClr = -1) Then
        CellForeColor lHdc, lOldColor
    End If
    '/* restore font
    If Not (lOldFont = 0) Then
        SelectObject lHdc, lOldFont
        lOldFont = 0
    End If

End Sub

Public Sub CellErase(ByVal lRow As Long, _
                     ByVal lCell As Long, _
                     ByVal bErase As Boolean)

'/* erase the cell

Dim lFlag  As Long
Dim tRect  As RECT

    If Not (m_lRowCount = 0) Then
        If (lRow < m_lRowCount) Then
            If bErase Then
                lFlag = 1
            Else
                lFlag = 0
            End If
            CellCalcRect lRow, lCell, tRect
            With tRect
                If Not (m_eGridLines = EGLNone) Then
                    InflateRect tRect, 0, -1
                    .top = .top - 1
                End If
                If Not m_bFullRowSelect Then
                    If m_bColumnFocus Then
                        If (lCell = CellHitTest) Then
                            InflateRect tRect, -2, 0
                            .left = .left - 1
                        End If
                    End If
                End If
            End With
            EraseRect m_lHGHwnd, tRect, lFlag
        End If
    End If

End Sub

Private Sub CellForeColor(ByVal lHdc As Long, _
                          ByVal lColor As Long, _
                          Optional ByVal lBaseClr As Long = -1)
'/* cell text color

    If Not m_bEnabled Then
        lColor = m_lDisabledForeColor
    Else
    If m_bForeColorAuto Then
        If Not (lBaseClr = -1) Then
            If ((2& * (lBaseClr And &HFF) + 5& * ((lBaseClr \ &H100&) And &HFF) + _
                ((lBaseClr \ &H10000) And &HFF)) > &H400) Then
                lColor = &H0
            Else
                lColor = &HFFFFFF
            End If
        Else
            If (lColor = -1) Then
                If (m_oForeColor = -1) Then
                    lColor = &H0
                Else
                    lColor = m_oForeColor
                End If
            End If
        End If
    Else
        If (lColor = -1) Then
            If (m_oForeColor = -1) Then
                lColor = &H0
            Else
                lColor = m_oForeColor
            End If
        End If
    End If
    End If
    SetTextColor lHdc, lColor

End Sub

Private Function CellHitTest() As Long
'/* cell hit test

Dim lRow    As Long
Dim tLVH    As LVHITTESTINFO
Dim tPoint  As POINTAPI

    If Not (m_lRowCount = 0) Then
        CellHitTest = -1
        '/* relative position
        GetCursorPos tPoint
        ScreenToClient m_lHGHwnd, tPoint
        With tLVH
            .pt.x = tPoint.x
            .pt.y = tPoint.y
            .flags = LVHT_ONITEM
        End With
        '/* hit test
        SendMessageA m_lHGHwnd, LVM_SUBITEMHITTEST, 0&, tLVH
        CellHitTest = tLVH.iSubItem
        lRow = RowHitTest
        '/* test vert span state
        If Not m_bVirtualMode Then
            If Not (lRow = -1) Then
                If Not (m_cGridItem(lRow).SpanFirstCell = -1) Then
                    If CellIsSpanned(lRow, ColumnIndex(CellHitTest)) Then
                        CellHitTest = m_cGridItem(lRow).SpanFirstCell
                    Else
                        CellHitTest = ColumnIndex(CellHitTest)
                    End If
                End If
            End If
        End If
    End If

End Function

Private Function CellIsSpanned(ByVal lRow As Long, _
                               ByVal lCell As Long) As Boolean
'/* cell is horizontally spanned

    With m_cGridItem(lRow)
        If Not (.SpanFirstCell = -1) Then
            If (lCell >= .SpanFirstCell) Then
                If (lCell <= .SpanLastCell) Then
                    CellIsSpanned = True
                End If
            End If
        End If
    End With

End Function

Private Sub CellRedraw()
'/* redraw cell on click

Dim lRow    As Long
Dim lCell   As Long
Dim tRect   As RECT

On Error Resume Next

    If Not (m_lRowCount = 0) Then
        lRow = RowHitTest
        If (lRow > -1) Then
            lCell = CellHitTest
            If m_bFullRowSelect Then
                CellCalcRect lRow, -1, tRect
            Else
                CellCalcRect lRow, lCell, tRect
            End If
            EraseRect m_lHGHwnd, tRect, 0&
            If (lRow = RowInFocus) Then
                If Not (lCell = CellInFocus) Then
                    If m_bFullRowSelect Then
                        CellCalcRect RowInFocus, -1, tRect
                    Else
                        CellCalcRect RowInFocus, CellInFocus, tRect
                    End If
                    EraseRect m_lHGHwnd, tRect, 0&
                End If
            Else
                If m_bFullRowSelect Then
                    CellCalcRect RowInFocus, -1, tRect
                Else
                    CellCalcRect RowInFocus, CellInFocus, tRect
                End If
                EraseRect m_lHGHwnd, tRect, 0&
            End If
            RowInFocus = lRow
            CellInFocus = lCell
        End If
        RaiseEvent eVHItemClick(lRow, lCell)
    End If

On Error GoTo 0

End Sub

Private Sub CellRefresh(ByVal lRow As Long, _
                        Optional ByVal lCell As Long = -1)
'/* refresh  a row

Dim tRect As RECT

    If (lRow > -1) Then
        If (lCell = -1) Then
            CellCalcRect lRow, -1, tRect
        Else
            CellCalcRect lRow, lCell, tRect
        End If
        EraseRect m_lHGHwnd, tRect, 0&
    End If

End Sub

Private Function CellSelected(ByRef tRect As RECT) As Boolean
'/* cell selected hit test

Dim tPoint As POINTAPI

    GetCursorPos tPoint
    ScreenToClient m_lHGHwnd, tPoint
    With tPoint
        If Not (PtInRect(tRect, .x, .y) = 0) Then
            CellSelected = True
        End If
    End With

End Function

Private Sub CellTipStart()
'/* load cell tip

    With m_cCellTips
        .CtrlHwnd = m_lHGHwnd
        .ImlHwnd = m_lImlRowHndl
        .XPColors = m_bCellTipXPColors
        .BackColor = m_oCellTipColor
        .ColorOffset = m_oCellTipOffsetColor
        .DelayTime = m_lCellTipDelayTime
        .ForeColor = m_oCellTipForeColor
        .Gradient = m_bCellTipGradient
        .Multiline = m_bCellTipMultiline
        .ToolTipPosition = m_lCellTipPosition
        .Transparency = m_lCellTipTransparency
        .VisibleTime = m_lCellTipVisibleTime
        .FontRightLeading = m_bFontRightLeading
        .Width = 120
        .UseUnicode = m_bUseUnicode
        If Not (m_oCellTipFont Is Nothing) Then
            Set .Font = m_oCellTipFont
        End If
    End With

End Sub

Private Sub CellTipTrack()
'/* track cell tip state

Dim lHdc    As Long
Dim lRow    As Long
Dim lCell   As Long
Dim lSLen   As Long
Dim lIcon   As Long
Dim lClHdr  As Long
Dim sText   As String
Dim sTitle  As String
Dim tPnt    As POINTAPI
Dim tRect   As RECT

    '/* in screen
    GetCursorPos tPnt
    ScreenToClient m_lHGHwnd, tPnt
    GetClientRect m_lHGHwnd, tRect
    tRect.top = m_lHeaderHeight + 2
    If (PtInRect(tRect, tPnt.x, tPnt.y) = 0) Then
        StopTipTimer
        Exit Sub
    End If
    
    '/* cell hit test
    lRow = RowHitTest
    lCell = CellHitTest
    If (lRow = -1) Then Exit Sub
    If (lCell = -1) Then Exit Sub
    
    If (lRow = m_lLastRow) Then
        If (lCell = m_lLastCell) Then
            Exit Sub
        Else
            StopTipTimer
        End If
    Else
        StopTipTimer
    End If
    m_lLastRow = lRow
    m_lLastCell = lCell
    
    If m_bFullRowSelect Then
        If RowFocused(lRow) Then
            Exit Sub
        End If
    Else
        If CellFocused(lRow, lCell) Then
            Exit Sub
        End If
    End If
    
    If m_bItemActive Then
        Exit Sub
    End If

    With m_cCellTips
        lIcon = CellIcon(lRow, lCell)
        sText = CellText(lRow, lCell)
        lClHdr = m_cGridItem(lRow).CellHeader(lCell)
        If (lClHdr > -1) Then
            sTitle = m_cCellHeader(lClHdr).Text
            If (LenB(sTitle) = 0) Then
                sTitle = ColumnText(lCell)
            End If
        Else
            sTitle = ColumnText(lCell)
        End If
        If Not (lIcon = -1) Then
            lSLen = m_lRowIconX
        End If
        lHdc = GetDC(m_lHGHwnd)
        If m_bIsNt Then
            If (lstrlenW(StrPtr(sText)) > 50) Then
                CharDimensions lHdc, tPnt
                lSLen = (tPnt.x * lstrlenW(StrPtr(sText))) / 4
                .Width = lSLen
            ElseIf (LenB(sText) < LenB(sTitle)) Then
                CharDimensions lHdc, tPnt
                lSLen = lSLen + (tPnt.x * lstrlenW(StrPtr(sTitle))) + 8
                .Width = lSLen
            End If
        Else
            If (Len(sText) > 50) Then
                CharDimensions lHdc, tPnt
                lSLen = (tPnt.x * Len(sText)) / 4
                .Width = lSLen
            ElseIf (Len(sText) < Len(sTitle)) Then
                CharDimensions lHdc, tPnt
                lSLen = lSLen + (tPnt.x * Len(sTitle)) + 8
                .Width = lSLen
            End If
        End If
        ReleaseDC m_lHGHwnd, lHdc
        .Icon = lIcon
        .Text = sText
        .Title = sTitle
    End With
    
    If Not m_bShowing Then
        If (CellHitTest > -1) Then
            If (RowHitTest > -1) Then
                StartTipTimer
            End If
        End If
    End If

End Sub

Private Sub CellTrack()
'/* track and highlight cells

Dim lRow       As Long
Dim lHdc       As Long
Dim lhBrush    As Long
Dim lCt        As Long
Dim tRect      As RECT
Dim tRCell     As RECT

    If Not (m_lRowCount = 0) Then
        lRow = RowHitTest
        If (lRow = -1) Then
            Exit Sub
        ElseIf (lRow = m_lLastRow) Then
            If m_bFullRowSelect Then
                Exit Sub
            End If
        End If
        CellCalcRect lRow, -1, tRect
        If m_bFullRowSelect Then
            If m_bFirstRowReserved Then
                CellCalcRect lRow, ColumnAtIndex(0), tRCell
                tRect.left = tRCell.Right
            End If
        End If
        If CellSelected(tRect) Then
            If m_bFullRowSelect Then
                CellErase m_lLastRow, -1, False
            Else
                lCt = CellHitTest
                If (m_lLastCell = lCt) And (m_lLastRow = lRow) Then
                    GoTo Handler
                Else
                    CellErase m_lLastRow, m_lLastCell, False
                    CellCalcRect lRow, lCt, tRect
                    m_lLastCell = lCt
                    With tRect
                        If m_bColumnFocus Then
                            InflateRect tRect, -2, 0
                            .left = .left - 1
                        Else
                            .Right = .Right - 1
                        End If
                    End With
                End If
            End If
            If Not (m_eGridLines = EGLNone) Then
                InflateRect tRect, 0, -1
                With tRect
                    .top = .top - 1
                End With
            End If
            If m_bColumnFocus Then
                If lCt = m_lColumnSelected Then
                    InflateRect tRect, -1, 0
                End If
            End If
            '/* draw frame
            lHdc = GetDC(m_lHGHwnd)
            lhBrush = CreateSolidBrush(m_oHotTrackColor)
            FrameRect lHdc, tRect, lhBrush
            If (m_eHotTrackDepth = ehtWide) Then
                OffsetRect tRect, -1, -1
                InflateRect tRect, -2, -2
                FrameRect lHdc, tRect, lhBrush
            End If
            ReleaseDC m_lHGHwnd, lHdc
            DeleteObject lhBrush
            m_lLastRow = lRow
        End If
    End If

Handler:
    m_lLastCell = lCt
    m_lLastRow = lRow

End Sub

Private Function CharDimensions(ByVal lHdc As Long, _
                                ByRef tPnt As POINTAPI)
'/* size of font in dc

Dim tTMtc As TEXTMETRIC

    GetTextMetricsA lHdc, tTMtc
    With tTMtc
        tPnt.y = .tmHeight + .tmExternalLeading + 1
        tPnt.x = .tmAveCharWidth
    End With

End Function

Private Function CheckBoxHitTest() As Boolean
'/* checkbox hit test

Dim lRow    As Long
Dim lCell   As Long
Dim tRect   As RECT
Dim tPnt    As POINTAPI

    If m_bCheckBoxes Then
        lCell = CellHitTest
        If (lCell = 0) Then
            lRow = RowHitTest
            If (lRow > -1) Then
                CellCalcRect lRow, lCell, tRect
                CalcCheckBoxRect tRect
                GetCursorPos tPnt
                ScreenToClient m_lHGHwnd, tPnt
                With tPnt
                    If Not (PtInRect(tRect, .x, .y) = 0) Then
                        If LeftKeyState Then
                            CheckToggle lRow
                            CheckBoxHitTest = True
                        End If
                    End If
                End With
            End If
        End If
    End If

End Function

Private Sub ClipNonClient()
'/* compensates for spanned row
'/* bleeding into header non-client

Dim lSdc    As Long
Dim tRSub   As RECT
Dim tRClt   As RECT

    If m_bEnabled Then
        GetClientRect m_lHGHwnd, tRSub
        With tRSub
            lSdc = GetDC(m_lHGHwnd)
            With tRClt
                .Right = tRSub.Right
                .top = m_lHeaderHeight
                .Bottom = m_lHeaderOffset - m_lHeaderHeight
            End With
            EraseRect m_lHGHwnd, tRClt, 0&
            ExcludeClipRect lSdc, .left, .top, .Right, m_lHeaderOffset
            ReleaseDC m_lHGHwnd, lSdc
        End With
    End If

End Sub

Private Sub CreateTransitionMask()
'/* create column vertical sizing client mask

Dim lHdc    As Long
Dim tRWnd   As RECT
Dim tRClt   As RECT

    If Not (m_lRowCount = 0) Then
        If (m_cTransitionMask Is Nothing) Then
            Set m_cTransitionMask = New clsStoreDc
        Else
            Exit Sub
        End If
        GetWindowRect m_lHGHwnd, tRWnd
        GetClientRect m_lHGHwnd, tRClt
        lHdc = GetDC(m_lHGHwnd)
        '/* create the dc
        With tRClt
            m_cTransitionMask.Width = .Right - .left
            m_cTransitionMask.Height = (.Bottom - .top) - m_lHeaderHeight
        End With
        '/* blit screen to dc
        With tRClt
            m_cRender.Blit m_cTransitionMask.hdc, 0, 0, (.Right - .left), (.Bottom - .top), lHdc, .left, (.top + m_lHeaderHeight), SRCCOPY
        End With
        ReleaseDC m_lHGHwnd, lHdc
    End If

End Sub

Private Sub DestroyTransitionMask()
'/* destroy transition mask

    If Not (m_cTransitionMask Is Nothing) Then
        Set m_cTransitionMask = Nothing
    End If
    m_lInitialHeaderHeight = 0
    m_bTransitionMask = False
    SetRowCount Count
    GridRefresh False

End Sub

Private Sub DrawTransitionMask(ByVal lHdc As Long)
'/* draw transition mask

Dim tRClt As RECT
    
    If Not (m_cTransitionMask Is Nothing) Then
        GetClientRect m_lHGHwnd, tRClt
        If (m_lInitialHeaderHeight = 0) Then
            m_lInitialHeaderHeight = m_lHeaderHeight + 2
        End If
        With m_cTransitionMask
            m_cRender.Blit lHdc, 0, m_lInitialHeaderHeight, tRClt.Right, .Height, .hdc, 0, 0, SRCCOPY
        End With
    End If
    
End Sub

Private Sub DrawCellAlpha(ByVal lHdc As Long, _
                          ByRef tRect As RECT)
'/* draw alpha bar

Dim lDrawDc As Long
Dim lBmp    As Long
Dim lBmpOld As Long
Dim tTmp    As RECT

    lDrawDc = CreateCompatibleDC(lHdc)
    CopyRect tTmp, tRect
    With tTmp
        OffsetRect tTmp, -.left, -.top
        lBmp = CreateCompatibleBitmap(lHdc, .Right, .Bottom)
    End With
    lBmpOld = SelectObject(lDrawDc, lBmp)
    '/* paint
    With tTmp
        '/* left
        m_cRender.Stretch lDrawDc, 0, 3, 3, (.Bottom - 6), m_cSelectorBar.hdc, 0, 3, 3, (m_cSelectorBar.Height - 6), SRCCOPY
        '/* top
        m_cRender.Stretch lDrawDc, 0, .top, .Right, 3, m_cSelectorBar.hdc, 3, 0, m_cSelectorBar.Width - 6, 3, SRCCOPY
        '/* bottom
        m_cRender.Stretch lDrawDc, 0, (.Bottom - 3), .Right, 3, m_cSelectorBar.hdc, 3, (m_cSelectorBar.Height - 3), (m_cSelectorBar.Width - 6), 3, SRCCOPY
        '/* center
        m_cRender.Stretch lDrawDc, 3, 3, (.Right - 6), (.Bottom - 6), m_cSelectorBar.hdc, 3, 3, (m_cSelectorBar.Width - 6), (m_cSelectorBar.Height - 6), SRCCOPY
        '/* right
        m_cRender.Stretch lDrawDc, (.Right - 3), 3, 3, (.Bottom - 6), m_cSelectorBar.hdc, (m_cSelectorBar.Width - 3), 3, 3, (m_cSelectorBar.Height - 6), SRCCOPY
        '/* copy to dest
        m_cRender.AlphaBlit lHdc, tRect.left, tRect.top, .Right, .Bottom, lDrawDc, 0, 0, .Right, .Bottom, m_bteAlphaTransparency
    End With
    '/* cleanup
    SelectObject lDrawDc, lBmpOld
    DeleteObject lBmp
    DeleteDC lDrawDc

End Sub

Public Property Get OwnerDrawImpl() As IOwnerDrawn
Attribute OwnerDrawImpl.VB_MemberFlags = "40"
'/* [get] ownerdrawn interface
    If Not (m_lPtrOwnerDraw = 0) Then
        Set OwnerDrawImpl = ObjectFromPtr(m_lPtrOwnerDraw)
    End If
End Property

Public Property Let OwnerDrawImpl(PropVal As IOwnerDrawn)
'/* [let] ownerdrawn interface
    SetOwnerDrawImpl PropVal
End Property

Public Property Set OwnerDrawImpl(PropVal As IOwnerDrawn)
'/* [set] ownerdrawn interface
    SetOwnerDrawImpl PropVal
End Property

Private Sub SetOwnerDrawImpl(PropVal As IOwnerDrawn)
'/* reference ownerdrawn interface
    If (PropVal Is Nothing) Then
        m_lPtrOwnerDraw = 0
    Else
        m_lPtrOwnerDraw = ObjPtr(PropVal)
    End If
End Sub

Public Property Get ObjectFromPtr(ByVal lPointer As Long) As Object
Attribute ObjectFromPtr.VB_MemberFlags = "40"
'/* [get] object from a pointer
Dim oTemp As Object

   CopyMemory oTemp, lPointer, 4&
   Set ObjectFromPtr = oTemp
   CopyMemory oTemp, 0&, 4&

End Property

Private Sub CalcTextRect(ByVal lRow As Long, _
                         ByVal lCell As Long, _
                         ByVal lHdc As Long, _
                         ByRef tRect As RECT)

Dim lAlign      As Long
Dim lOldFont    As Long
Dim lRight      As Long
Dim lDist       As Long
Dim sText       As String
Dim tRTxt       As RECT

    CopyRect tRTxt, tRect
    sText = m_cGridItem(lRow).Text(lCell)
    If LenB(sText) > 0 Then
        If Not (m_cGridItem(lRow).FontHnd(lCell) = 0) Then
            lOldFont = SelectObject(lHdc, m_cGridItem(lRow).FontHnd(lCell))
        End If
        sText = sText & vbNullChar
        lAlign = m_cGridItem(lRow).Align(lCell)
        lRight = tRect.Right
        '/* get rect size
        If m_bIsNt Then
            DrawTextW lHdc, StrPtr(sText), -1, tRTxt, lAlign Or DT_CALCRECT
        Else
            DrawTextA lHdc, sText, -1, tRTxt, lAlign Or DT_CALCRECT
        End If
        If Not (lOldFont = 0) Then
            SelectObject lHdc, lOldFont
            lOldFont = 0
        End If
        If tRect.Right > lRight Then
            tRect.Right = lRight
        End If
    End If
    '/*
    With tRTxt
        If (.Right > tRect.Right) Then
            .Right = tRect.Right
        End If
        If (.Bottom > tRect.Bottom) Then
            .Bottom = tRect.Bottom
        End If
        If (lAlign And DT_VCENTER) = DT_VCENTER Then
            lDist = (tRect.Bottom - .Bottom) / 2
            If (.Bottom > tRect.Bottom) Then
                lDist = lDist - (tRTxt.Bottom - (tRect.Bottom + 2))
            End If
            OffsetRect tRTxt, 0&, lDist
        ElseIf (lAlign And DT_BOTTOM) = DT_BOTTOM Then
            lDist = tRect.Bottom - .Bottom
            OffsetRect tRTxt, 0&, lDist
        End If
        If (lAlign And DT_RIGHT) = DT_RIGHT Then
            lDist = tRect.Right - .Right
            OffsetRect tRTxt, lDist, 0&
        ElseIf (lAlign And DT_CENTER) = DT_CENTER Then
            lDist = (tRect.Right - .Right) / 2
            OffsetRect tRTxt, lDist, 0&
        End If
    End With
    CopyRect tRect, tRTxt

End Sub

Private Sub DrawGrid(ByVal lHdc As Long, _
                     ByVal lRow As Long)

'/* central cell drawing hub

Dim bFocus          As Boolean
Dim bSelect         As Boolean
Dim bGhosted        As Boolean
Dim bText           As Boolean
Dim bHideCheck      As Boolean
Dim bSkipDefault    As Boolean
Dim lCell           As Long
Dim lCount          As Long
Dim lAlClr          As Long
Dim lhBrush         As Long
Dim lFcsClr         As Long
Dim lRPtr           As Long
Dim lIcon           As Long
Dim lIndent         As Long
Dim lForeColor      As Long
Dim lBackColor      As Long
Dim lBfDc           As Long
Dim lAlign          As Long
Dim sText           As String
Dim tRcpy           As RECT
Dim tRSub           As RECT
Dim tRTxt           As RECT
Dim tRFcs           As RECT
Dim tRIcn           As RECT
Dim tRect           As RECT
Dim IODraw          As IOwnerDrawn

    If Not (m_lPtrOwnerDraw = 0) Then
        Set IODraw = ObjectFromPtr(m_lPtrOwnerDraw)
    End If

    If m_bUseSpannedRows Then
        lRow = RowSpanMapVirtual(lRow)
    End If
    
    '/* store base rect
    CellCalcRect lRow, -1, tRect
    '/* nc clipping of spanned rows
    ClipNonClient
    '/* store focus rect
    CopyRect tRFcs, tRect
    
    '/* draw row into buffer
    If m_bDoubleBuffer Then
        With m_cGridBuffer
            .Height = tRect.Bottom
            .Width = tRect.Right
            lBfDc = lHdc
            lHdc = .hdc
            SelectObject m_cGridBuffer.hdc, m_lFont
        End With
    End If

    lCount = m_lColumnCount - 1
    If RowFocused(lRow) Then
        If m_bGridFocused Then
            bFocus = True
        Else
            bSelect = True
        End If
    End If

    If m_bItemActive Then
        bSelect = False
        bFocus = False
    Else
        If Not m_bVirtualMode Then
            If m_cGridItem(lRow).RowNoFocus Then
                bSelect = False
                bFocus = False
            End If
        End If
    End If
    
    lAlClr = m_cRender.BlendColor(m_oCellFocusedColor, &HFFFFFF, 92)
    bGhosted = RowGhosted(lRow)
    If Not m_bVirtualMode Then
        bHideCheck = m_cGridItem(lRow).HideCheckBox
    End If
    
    '/* draw each grid cell
    For lCell = 0 To lCount

        '/*** get base rect ***/
        CellCalcRect lRow, lCell, tRSub
        '/* store base rect
        CopyRect tRcpy, tRSub
        '/* copy icon rect
        CopyRect tRIcn, tRSub

        '/* owner drawn pre-draw stage
        With tRSub
            If Not (m_lPtrOwnerDraw = 0) Then
                IODraw.Draw m_cGridItem(lRow), lRow, lCell, lHdc, edgPreDraw, .left, .top, .Right, .Bottom, bSkipDefault
            End If
        End With
        
        '/*** store class variables ***/
        If m_bVirtualMode Then
            RaiseEvent eVHVirtualAccess(lRow, lCell, sText, lIcon)
        Else
            With m_cGridItem(lRow)
                lIcon = .Icon(lCell)
                lIndent = .Indent(lCell)
                lAlign = .Align(lCell)
                lBackColor = .BackColor(lCell)
                lForeColor = .ForeColor(lCell)
                sText = .Text(lCell)
            End With
        End If
        
        '/*** cell pattern colors ***/
        If m_bCellDecoration Then
            lBackColor = CellColor(lRow, lCell)
        End If
        
        '/*** draw grid lines ***/
        With tRSub
            '/* both horz and vert lines
            If (m_eGridLines = EGLBoth) Then
                '/* horz grid line
                .Bottom = .Bottom - 1
                GridLine lHdc, .left, .Bottom, .Right, .Bottom, m_oGridLineColor, 1
                '/* vert grid line
                .Right = .Right - 1
                If (lCell = m_lColumnDivider) Then
                    GridLine lHdc, .Right, .top, .Right, .Bottom - 1, &H4422EE, 1
                    GridLine lHdc, .Right - 1, .top, .Right - 1, .Bottom - 1, &H4422EE, 1
                    .Right = .Right - 1
                Else
                    GridLine lHdc, .Right, .top, .Right, .Bottom - 1, m_oGridLineColor, 1
                End If
            '/* horz lines
            ElseIf (m_eGridLines = EGLHorizontal) Then
                .Bottom = .Bottom - 1
                GridLine lHdc, .left, .Bottom, .Right, .Bottom, m_oGridLineColor, 1
                If lRow = 0 Then
                    GridLine lHdc, .left, .top, .Right, .top, m_oGridLineColor, 1
                    .top = .top + 1
                End If
            '/* vert lines
            ElseIf (m_eGridLines = EGLVertical) Then
                .Right = .Right - 1
                GridLine lHdc, .Right, .top, .Right, .Bottom, m_oGridLineColor, 1
            End If
        End With
        
        '/* owner drawn before background
        With tRSub
            If Not (m_lPtrOwnerDraw = 0) Then
                IODraw.Draw m_cGridItem(lRow), lRow, lCell, lHdc, edgBeforeBackGround, .left, .top, .Right, .Bottom, bSkipDefault
            End If
        End With
        
        '/*** column focus ***/
        If Not m_bFullRowSelect Then
            If m_bColumnFocus Then
                If (lCell = CellInFocus) Then
                    With tRSub
                        GridLine lHdc, .left, .top, .left, .Bottom, m_oColumnFocusColor, 1
                        GridLine lHdc, .Right - 1, .top, .Right - 1, .Bottom, m_oColumnFocusColor, 1
                        InflateRect tRSub, -1, 0
                        If lBackColor > &HFFFFF Then
                            lBackColor = lBackColor - &HF0F0F
                        Else
                            lBackColor = lBackColor Or &H8
                        End If
                    End With
                End If
            End If
        End If
        
        '/*** checkboxes ***/
        If (lCell = 0) Then
            If m_bCheckBoxes Then
                If Not bHideCheck Then
                    CellDrawCheckBox lRow, lHdc, tRSub, bGhosted
                    CopyRect tRcpy, tRSub
                    CalcCheckBoxRect tRcpy
                    With tRcpy
                        ExcludeClipRect lHdc, .left, .top, .Right, .Bottom
                    End With
                    If (lIcon > -1) Then
                        With tRIcn
                            .left = .left + 18
                        End With
                    End If
                End If
            End If
        End If
        
        '/*** sort pointer ***/
        If m_bFiltered Then
            lRPtr = m_lFilter(lRow)
        Else
            lRPtr = lRow
        End If
        
        If Not bSkipDefault Then
            bText = Not (LenB(sText) = 0)
            If bText Then
                CopyRect tRTxt, tRSub
                '/* calculate initial offsets
                InflateRect tRTxt, -2, -1
                With tRTxt
                    If (lIndent > 0) Then
                        .left = (.left + lIndent)
                    End If
                    If (lIcon > -1) Then
                        .left = (.left + m_lRowIconX + 2)
                    End If
                    If (lCell = 0) Then
                        If m_bCheckBoxes Then
                            .left = (.left + 18)
                        End If
                    End If
                End With
            End If

            '/*** cell backcolor ***/
            If bGhosted Then
                CellBackColor lHdc, &HDCDCDC, tRSub
            ElseIf bFocus Then
                '/* color selection
                If m_bAlphaBlend Then
                    lFcsClr = lAlClr
                Else
                    lFcsClr = m_oCellFocusedColor
                End If
                '/* first cell fixed
                If m_bFirstRowReserved Then
                    If lCell = 0 Then
                        lFcsClr = lBackColor
                    End If
                End If
            
                '/*** full row focused ***/
                If m_bFullRowSelect Then
                    If Not m_bAlphaBlend Then
                        '/* offset for focused rect
                        If Not m_bAlphaSelectorBar Then
                            With tRSub
                                If lCell = ColumnIndex(lCount) Then
                                    .Right = .Right - 1
                                Else
                                    If m_bFirstRowReserved Then
                                        If lCell = ColumnIndex(1) Then
                                            .left = .left + 1
                                        End If
                                    Else
                                        If lCell = ColumnIndex(0) Then
                                            .left = .left + 1
                                        End If
                                    End If
                                End If
                            End With
                            InflateRect tRSub, 0, -1
                        End If
                    End If
                    '/* exclude alphabar
                    If m_bAlphaSelectorBar Then
                        CellBackColor lHdc, lBackColor, tRSub
                    Else
                        CellBackColor lHdc, lFcsClr, tRSub
                    End If
            
                '/*** per cell focused ***/
                Else
                    If CellFocused(lRow, lCell) Then
                        If m_bAlphaSelectorBar Then
                            CellBackColor lHdc, lBackColor, tRSub
                            DrawCellAlpha lHdc, tRSub
                        Else
                            '/* cell text only focus
                            If m_bFocusTextOnly Then
                                If bText Then
                                    '/* calculate focused rect
                                    CopyRect tRcpy, tRTxt
                                    On Error Resume Next
                                    CalcTextRect lRow, lCell, lHdc, tRcpy
                                    If m_cGridItem(lRow).CellHeader(lCell) > -1 Then
                                        CellCalcHeaderSize lRow, lCell, lHdc, sText, tRcpy
                                        '*********************************************
                                      '  OffsetRect tRcpy, 0, -5
                                    End If
                                    On Error GoTo 0
                                    InflateRect tRcpy, 4, 2
                                    OffsetRect tRcpy, 4, -1
                                    '/* back color draw
                                    With tRcpy
                                        If Not (.Right < tRSub.Right) Then
                                            .Right = tRSub.Right
                                        End If
                                        If .top = tRSub.top Then
                                            .top = .top + 1
                                        End If
                                        If (.Bottom > tRSub.Bottom) Then
                                            .Bottom = (tRSub.Bottom - 2)
                                        End If
                                    End With
                                    
                                    CellBackColor lHdc, lBackColor, tRSub
                                    '/* deselect region
                                    SelectClipRgn lHdc, 0&
                                    '/* draw focus rect
                                    DrawFocusRect lHdc, tRcpy
                                    InflateRect tRcpy, -1, -1
                                    '/* fill rect w/ focus color
                                    CellBackColor lHdc, lFcsClr, tRcpy
                                End If
                            '/* standard focus
                            Else
                                CellBackColor lHdc, lFcsClr, tRSub
                            End If
                        End If
                    '/* not focused
                    Else
                        CellBackColor lHdc, lBackColor, tRSub
                    End If
                End If

            '/*** selected only backcolor ***/
            ElseIf bSelect Then
                If m_bFullRowSelect Then
                    If RowFocused(lRow) Then
                        CellBackColor lHdc, m_oCellSelectedColor, tRSub
                    Else
                        CellBackColor lHdc, lBackColor, tRSub
                    End If
                Else
                    If CellFocused(lRow, lCell) Then
                        If lCell = 0 Then
                            If m_bFirstRowReserved Then
                                CellBackColor lHdc, lBackColor, tRSub
                            Else
                                CellBackColor lHdc, m_oCellSelectedColor, tRSub
                            End If
                        Else
                            CellBackColor lHdc, m_oCellSelectedColor, tRSub
                        End If
                    Else
                        CellBackColor lHdc, lBackColor, tRSub
                    End If
                End If
            '/*** normal backcolor ***/
            Else
                CellBackColor lHdc, lBackColor, tRSub
            End If
        End If

        '/*** store pointer offsets ***/
        If Not m_bVirtualMode Then
            If Not (lRPtr = lRow) Then
                With m_cGridItem(lRPtr)
                    lIcon = .Icon(lCell)
                    lIndent = .Indent(lCell)
                    sText = .Text(lCell)
                End With
            End If
        End If
        
        '/* owner drawn pre icon
        With tRSub
            If Not (m_lPtrOwnerDraw = 0) Then
                IODraw.Draw m_cGridItem(lRow), lRow, lCell, lHdc, edgBeforeIcon, .left, .top, .Right, .Bottom, bSkipDefault
            End If
        End With
        
        If Not bSkipDefault Then
            '/*** icon focus ***/
            If (lIcon > -1) Then
                '/* first cell special
                If Not (lCell = 0) Then
                    '/* apply cell indenting to icon
                    If (lIndent > 0) Then
                        With tRIcn
                            .left = .left + lIndent
                        End With
                    End If
                End If
                '/* cell icon focus effects
                If m_bFullRowSelect Then
                    If (lCell = 0) Then
                        If bGhosted Then
                            CellDrawIcon lHdc, lIcon, 2, tRIcn
                        ElseIf bFocus Then
                            '/* focus fade
                            If m_bAlphaBlend Then
                                CellDrawIcon lHdc, lIcon, 3, tRIcn
                            '/* normal
                            Else
                                CellDrawIcon lHdc, lIcon, 2, tRIcn
                            End If
                        '/* selected fade
                        ElseIf bSelect Then
                            CellDrawIcon lHdc, lIcon, 1, tRIcn
                        '/* normal render
                        Else
                            CellDrawIcon lHdc, lIcon, 0, tRIcn
                        End If
                    Else
                        If bGhosted Then
                            CellDrawIcon lHdc, lIcon, 2, tRIcn
                        ElseIf bFocus Then
                            CellDrawIcon lHdc, lIcon, 3, tRIcn
                        Else
                            CellDrawIcon lHdc, lIcon, 0, tRIcn
                        End If
                    End If
                Else
                    If bGhosted Then
                        CellDrawIcon lHdc, lIcon, 2, tRIcn
                    ElseIf bFocus Then
                        If CellFocused(lRow, lCell) Then
                            CellDrawIcon lHdc, lIcon, 3, tRIcn
                        Else
                            CellDrawIcon lHdc, lIcon, 0, tRIcn
                        End If
                    Else
                        CellDrawIcon lHdc, lIcon, 0, tRIcn
                    End If
                End If
            Else
                SelectClipRgn lHdc, 0&
            End If
        End If
        
        '/* ownerdrawn pre text
        With tRSub
            If Not (m_lPtrOwnerDraw = 0) Then
                IODraw.Draw m_cGridItem(lRow), lRow, lCell, lHdc, edgBeforeText, .left, .top, .Right, .Bottom, bSkipDefault
            End If
        End With
        
        If Not bSkipDefault Then
            '/*** font forecolor ***/
            If bText Then
                If bGhosted Then
                    CellForeColor lHdc, &H808080
                ElseIf bFocus Then
                    '/* focused highlite
                    If m_bFullRowSelect Then
                        If m_bAlphaBlend Then
                            CellForeColor lHdc, m_oCellFocusedHighlight
                        '/* focused: exclude alpha bar
                        ElseIf Not m_bAlphaSelectorBar Then
                            CellForeColor lHdc, m_oCellFocusedHighlight, lFcsClr
                        Else
                            CellForeColor lHdc, lForeColor
                        End If
                    Else
                        If CellFocused(lRow, lCell) Then
                            If m_bAlphaBlend Then
                                CellForeColor lHdc, m_oCellFocusedHighlight
                            ElseIf Not m_bAlphaSelectorBar Then
                                CellForeColor lHdc, m_oCellFocusedHighlight, lFcsClr
                            End If
                        Else
                            CellForeColor lHdc, lForeColor
                        End If
                    End If
                '/* normal forecolor
                Else
                    CellForeColor lHdc, lForeColor
                End If
                InflateRect tRTxt, -1, -1
                OffsetRect tRTxt, 1, -1

                '/*** draw cell text ***/
                If bText Then
                    If Not m_bVirtualMode Then
                        If CellIsSpanned(lRow, lCell) Then
                            If lCell = CellSpanFirstCell(lRow) Then
                                CellDrawText lRPtr, lCell, lHdc, sText, tRTxt
                            End If
                        Else
                            CellDrawText lRPtr, lCell, lHdc, sText, tRTxt
                        End If
                    Else
                        CellDrawText lRPtr, lCell, lHdc, sText, tRTxt
                    End If
                End If
                '/* store edit rect
                If CellFocused(lRow, lCell) Then
                    CopyRect m_tREditCrd, tRTxt
                End If
            End If
        End If
        
        '/* vertically spanned
        If Not m_bVirtualMode Then
            If m_bFiltered Then
                If (lCell = m_cGridItem(m_lFilter(lRow)).SpanFirstCell) Then
                    DrawSpannedHeaders lHdc, lRow, lCell, m_cGridItem(m_lFilter(lRow)).SpanLastCell
                    lCell = m_cGridItem(m_lFilter(lRow)).SpanLastCell
                End If
            Else
                If (lCell = m_cGridItem(lRow).SpanFirstCell) Then
                    DrawSpannedHeaders lHdc, lRow, lCell, m_cGridItem(lRow).SpanLastCell
                    lCell = m_cGridItem(lRow).SpanLastCell
                End If
            End If
        End If
        
        '/* owner drawn post draw
        With tRSub
            If Not (m_lPtrOwnerDraw = 0) Then
                IODraw.Draw m_cGridItem(lRow), lRow, lCell, lHdc, edgPostDraw, .left, .top, .Right, .Bottom, bSkipDefault
            End If
        End With
        '/* reset for next cell
        bSkipDefault = False
    Next lCell

    '/*** full row focus ***/
    If Not bSkipDefault Then
        If Not bGhosted Then
            If bFocus Then
                If m_bFullRowSelect Then
                    If m_bFirstRowReserved Then
                        CellCalcRect lRow, ColumnAtIndex(0), tRcpy
                        tRFcs.left = tRcpy.Right - 1
                    End If
                    InflateRect tRFcs, -1, 0
                    tRFcs.Bottom = tRFcs.Bottom - 1
                    '/* alpha blended
                    If m_bAlphaBlend Then
                        lhBrush = CreateSolidBrush(TranslateColor(m_oCellFocusedColor))
                        FrameRect lHdc, tRFcs, lhBrush
                        DeleteObject lhBrush
                    '/* with alpha bar
                    ElseIf m_bAlphaSelectorBar Then
                        DrawCellAlpha lHdc, tRFcs
                    '/* solid rect
                    Else
                        lhBrush = CreateSolidBrush(TranslateColor(&H80808))
                        FrameRect lHdc, tRFcs, lhBrush
                        DeleteObject lhBrush
                    End If
                End If
            End If
        End If
    End If
    
    '/* position controls
    If m_bHasSubCells Then
        SubCellScrollCntl lHdc
    End If
    '/* header offset
    If tRSub.top < m_lHeaderHeight + 4 Then
        If Not m_bHeaderHide Then
            tRect.top = m_lHeaderHeight + 4
        End If
    End If
    '/* draw dc
    If m_bDoubleBuffer Then
        With tRect
            m_cRender.Blit lBfDc, .left, .top, (.Right - .left), (.Bottom - .top), lHdc, .left, .top, SRCCOPY
        End With
    End If

End Sub

Private Sub DrawSpannedHeaders(ByVal lHdc As Long, _
                               ByVal lRow As Long, _
                               ByVal lFirstCell As Long, _
                               ByVal lLastCell As Long)

Dim lCt     As Long
Dim lClHdr  As Long
Dim tRHdr   As RECT

    For lCt = lFirstCell To lLastCell
        lClHdr = m_cGridItem(lRow).CellHeader(lCt)
        If (lClHdr > -1) Then
            If (LenB(m_cCellHeader(lCt).Text) > 0) Then
                If (m_cGridItem(lRow).Icon(lCt) > -1) Then
                    CellCalcRect lRow, lCt, tRHdr
                End If
                CellDrawHeader lClHdr, lHdc, tRHdr.left + m_cCellHeader(lCt).Indent, lRow, lCt, tRHdr
            End If
        End If
    Next lCt
                
End Sub

Private Sub RowArrayResize()
'/* load vertical span tracking array

Dim lUb As Long
Dim lRc As Long

    If ArrayCheck(m_cGridItem) Then
        lUb = UBound(m_cGridItem)
        CalcRowOffsets
        lRc = (m_lRowDepth(UBound(m_lRowDepth)) / m_lRowHeight) - 1
        If lRc > lUb Then
            lUb = lRc
            m_bUseSpannedRows = True
        Else
            m_bUseSpannedRows = False
        End If
        If (lUb > 0) Then
            SendMessageA m_lHGHwnd, LVM_SETITEMCOUNT, (lUb + 1), LVSICF_NOINVALIDATEALL
        End If
    End If
    
End Sub

'**********************************************************************
'*                            FILTERS
'**********************************************************************

Public Property Get FilterInvert() As Boolean
Attribute FilterInvert.VB_MemberFlags = "40"
    FilterInvert = m_bFilterInvert
End Property

Public Property Let FilterInvert(ByVal PropVal As Boolean)
    m_bFilterInvert = PropVal
End Property

Public Property Get FilterHeader() As Boolean
Attribute FilterHeader.VB_MemberFlags = "40"
    FilterHeader = m_bFilterHeader
End Property

Public Property Let FilterHeader(ByVal PropVal As Boolean)
    m_bFilterHeader = PropVal
End Property

Public Property Get FilterHideExact() As Boolean
Attribute FilterHideExact.VB_MemberFlags = "40"
    FilterHideExact = m_bFilterHideExact
End Property

Public Property Let FilterHideExact(ByVal PropVal As Boolean)
    m_bFilterHideExact = PropVal
End Property

Public Sub FilterAdd(ByVal lColumn As Long, _
                     ByVal sFilter As String)
'/* add a filter item

    If Not m_bVirtualMode Then
        If (LenB(sFilter) = 0) Then Exit Sub
        FilterListAdd lColumn, sFilter
        ColumnFilters = True
        m_cSkinHeader.ColumnFiltered(lColumn) = True
    End If

End Sub

Private Sub FilterApply(ByVal lIndex As Long)
'/* apply a filter

Dim bMatch  As Boolean
Dim lLb     As Long
Dim lUb     As Long
Dim lCt     As Long
Dim lMatch  As Long
Dim sFilt   As String
Dim sArr()  As String

On Error GoTo Handler

    If Not (m_lRowCount = 0) Then
        lIndex = lIndex - 1
        sFilt = m_sFilterItem(m_lColumnFilter)
        If (InStr(1, m_sFilterItem(m_lColumnFilter), "|") > 0) Then
            sArr = Split(m_sFilterItem(m_lColumnFilter), "|")
            sFilt = sArr(lIndex)
        Else
            If (LenB(m_sFilterItem(m_lColumnFilter)) > 0) Then
                sFilt = m_sFilterItem(m_lColumnFilter)
            End If
        End If
    
        If (LenB(sFilt) = 0) Then
            GoTo Handler
        End If
    
        lLb = LBound(m_cGridItem)
        lUb = UBound(m_cGridItem)
        ReDim m_lFilter(lLb To lUb)
        lCt = lLb
        If Not m_bFilterHideExact Then
            If m_cFilterMenu.Checked Then
                bMatch = True
            End If
        End If
        If bMatch Then
            Do Until (lCt > lUb)
                If m_bFilterInvert Then
                    If (LenB(m_cGridItem(lCt).Text(m_lColumnFilter)) = LenB(sFilt)) Then
                        If (StrComp(m_cGridItem(lCt).Text(m_lColumnFilter), sFilt, vbBinaryCompare) = 0) Then
                            m_lFilter(lMatch) = lCt
                            lMatch = (lMatch + 1)
                        End If
                    End If
                Else
                    If Not (LenB(m_cGridItem(lCt).Text(m_lColumnFilter)) = LenB(sFilt)) Then
                        If Not (StrComp(m_cGridItem(lCt).Text(m_lColumnFilter), sFilt, vbBinaryCompare) = 0) Then
                            m_lFilter(lMatch) = lCt
                            lMatch = (lMatch + 1)
                        End If
                    End If
                End If
                lCt = lCt + 1
            Loop
        Else
            Do Until (lCt > lUb)
                If m_bFilterInvert Then
                    If (InStr(1, m_cGridItem(lCt).Text(m_lColumnFilter), sFilt, vbTextCompare) > 0) Then
                        m_lFilter(lMatch) = lCt
                        lMatch = (lMatch + 1)
                    End If
                Else
                    If Not (InStr(1, m_cGridItem(lCt).Text(m_lColumnFilter), sFilt, vbTextCompare) > 0) Then
                        m_lFilter(lMatch) = lCt
                        lMatch = (lMatch + 1)
                    End If
                End If
                lCt = lCt + 1
            Loop
        End If
        ReDim Preserve m_lFilter(0 To (lMatch - 1))
        If m_bUseSpannedRows Then
            For lCt = 0 To (lMatch - 1)
                If Not (m_cGridItem(lCt).SpanRowDepth = -1) Then
                    lMatch = lMatch + (m_cGridItem(lCt).SpanRowDepth - 1)
                End If
            Next lCt
        End If
        SetRowCount lMatch
        m_cSkinScrollBars.Refresh
    End If

Handler:
    On Error GoTo 0

End Sub

Private Sub FilterApplyHeader(ByVal lIndex As Long)
'/* apply a filter

Dim lLb     As Long
Dim lUb     As Long
Dim lCt     As Long
Dim lMatch  As Long
Dim sFilt   As String
Dim sArr()  As String

On Error GoTo Handler

    If Not (m_lRowCount = 0) Then
        lIndex = lIndex - 1
        sFilt = m_sFilterItem(m_lColumnFilter)
        If (InStr(1, m_sFilterItem(m_lColumnFilter), "|") > 0) Then
            sArr = Split(m_sFilterItem(m_lColumnFilter), "|")
            sFilt = sArr(lIndex)
        Else
            If (LenB(m_sFilterItem(m_lColumnFilter)) > 0) Then
                sFilt = m_sFilterItem(m_lColumnFilter)
            End If
        End If
    
        If (LenB(sFilt) = 0) Then
            GoTo Handler
        End If
    
        lLb = LBound(m_cCellHeader)
        lUb = UBound(m_cCellHeader)
        ReDim m_lFilter(lLb To lUb)

        lCt = lLb
        Do Until (lCt > lUb)
            If m_bFilterInvert Then
                If (InStr(1, m_cCellHeader(lCt).Text, sFilt, vbTextCompare) > 0) Then
                    m_lFilter(lMatch) = lCt
                    lMatch = (lMatch + 1)
                End If
            Else
                If Not (InStr(1, m_cCellHeader(lCt).Text, sFilt, vbTextCompare) > 0) Then
                    m_lFilter(lMatch) = lCt
                    lMatch = (lMatch + 1)
                End If
            End If
            lCt = lCt + 1
        Loop
        ReDim Preserve m_lFilter(0 To (lMatch - 1))
        SetRowCount lMatch
        m_cSkinScrollBars.Refresh
    End If

Handler:
    On Error GoTo 0

End Sub

Private Sub FilterItems(ByVal lColumn As Long)
'/* get filter list items

Dim lCt     As Long
Dim sArr()  As String

    If Not ArrayCheck(m_sFilterItem) Then
        Exit Sub
    ElseIf (UBound(m_sFilterItem) < lColumn) Then
        Exit Sub
    End If
    
    With m_cFilterMenu
        If (InStr(1, m_sFilterItem(lColumn), "|") > 0) Then
            sArr = Split(m_sFilterItem(lColumn), "|")
            For lCt = 0 To (UBound(sArr))
                If (LenB(sArr(lCt)) > 0) Then
                    .AddItem (lCt + 1), sArr(lCt)
                End If
            Next lCt
        Else
            If (LenB(m_sFilterItem(lColumn)) > 0) Then
                .AddItem 1, m_sFilterItem(lColumn)
            End If
        End If
    End With

End Sub

Public Sub FilterListAdd(ByVal lColumn As Long, _
                         ByVal sItem As String)
'/* add an item to filter list

    If (UBound(m_sFilterItem) < lColumn) Then
        ReDim Preserve m_sFilterItem(0 To lColumn)
    End If
    m_sFilterItem(lColumn) = sItem

End Sub

Public Sub FilterListRemove(ByVal lIndex As Long)
'/* remove an item from filter list

    If Not (lIndex = UBound(m_sFilterItem)) Then
        m_sFilterItem(lIndex) = UBound(m_sFilterItem)
    End If
    ReDim Preserve m_sFilterItem(0 To UBound(m_sFilterItem) - 1)
    If (UBound(m_sFilterItem) = 0) Then
        m_cSkinHeader.ColumnFiltered(lIndex) = False
    End If

End Sub

Private Sub FilterCreate(ByVal lColumn As Long)
'/* create the filter client

Dim lParHnd     As Long
Dim tPnt        As POINTAPI
Dim tPcd        As POINTAPI
Dim tRHdr       As RECT

    m_lColumnFilter = lColumn
    lParHnd = GetParent(m_lParentHwnd)
    GetWindowRect m_lHdrHwnd, tRHdr
    CopyMemory tPcd, tRHdr, Len(tPcd)
    ScreenToClient lParHnd, tPcd
    
    If Not (m_cFilterMenu Is Nothing) Then
        Set m_cFilterMenu = Nothing
    End If
    
    Set m_cFilterMenu = New clsFilterMenu
    GetCursorPos tPnt
    ScreenToClient m_lHdrHwnd, tPnt
    
    With tPnt
        .x = .x + tPcd.x
        .y = .y + tPcd.y
    End With
    
    With m_cFilterMenu
        .FilterHideExact = m_bFilterHideExact
        .Gradient = m_bFilterGradient
        .XPColors = m_bFilterXPColors
        .BackColor = m_oFilterBackColor
        .ColorOffset = m_oFilterOffsetColor
        .ControlColor = m_oFilterControlColor
        .ControlForeColor = m_oFilterControlForeColor
        .ForeColor = m_oFilterForeColor
        .TitleColor = m_oFilterTitleColor
        .Transparency = m_lFilterTransparency
        .AddItem 0, "None"
        .UseUnicode = m_bUseUnicode
        .FontRightLeading = m_bFontRightLeading
        FilterItems lColumn
        .ShowMenu m_lParentHwnd, "Filter Menu", tPnt.x, tPnt.y, 150, 120
    End With

End Sub

Private Sub FilterReset()
'/* reset filter

    ReDim m_lFilter(0)
    If m_bUseSpannedRows Then
        RowArrayResize
    Else
        SetRowCount m_lRowCount
    End If
    m_cSkinScrollBars.Refresh
    m_bSorted = False

End Sub

'>>>>>>>>>>>>>>>>>>>>>>>>>>
Public Sub FindStop()
    m_bStopSearch = True
End Sub

Public Function Find(ByVal sText As String, _
                     ByVal lColumn As Long, _
                     ByVal bMatchCase As Boolean, _
                     ByVal bExact As Boolean, _
                     ByVal bDescending As Boolean, _
                     ByVal bFindNext As Boolean, _
                     ByVal bDisplay As Boolean) As Long

Dim lCt     As Long
Dim lUt     As Long
Dim lRs     As Long
Dim lCm     As Long
Dim lSp     As Long

    If (m_lHGHwnd = 0) Then Exit Function
    If (RowCount = 0) Then Exit Function
    If m_bVirtualMode Then Exit Function
    
    lRs = -1
    lUt = (RowCount - 1)
    '/* shouldn't be putting more items then this
    If (lUt > 10000) Then Exit Function
    If (lColumn > (m_lColumnCount - 1)) Then Exit Function
    '/* store last found item
    If Not bFindNext Then
        lSp = 0
    Else
        If bDescending Then
            lSp = 0
        Else
            lSp = lSp + 1
        End If
    End If
    
    If bMatchCase Then
        lCm = vbBinaryCompare
    Else
        lCm = vbTextCompare
    End If
    
    If bDescending Then
        If Not (m_lRowFocused = 0) Then
            lUt = m_lRowFocused - 1
        Else
            lUt = UBound(m_cGridItem)
        End If
        For lCt = lUt To lSp Step -1
            '/* early exit
            If m_bStopSearch Then GoTo Handler
            If bExact Then
                If (sText = m_cGridItem(lCt).Text(lColumn)) Then
                    lRs = lCt
                    Exit For
                End If
            Else
                If (InStr(1, m_cGridItem(lCt).Text(lColumn), sText, lCm) > 0) Then
                    lRs = lCt
                    Exit For
                End If
            End If
        Next lCt
    Else
        If Not (m_lRowFocused = 0) Then
            lSp = m_lRowFocused + 1
        Else
            lSp = 0
        End If
        For lCt = lSp To lUt
            '/* early exit
            If m_bStopSearch Then GoTo Handler
            If bExact Then
                If (sText = m_cGridItem(lCt).Text(lColumn)) Then
                    lRs = lCt
                    Exit For
                End If
            Else
                If (InStr(1, m_cGridItem(lCt).Text(lColumn), sText, lCm) > 0) Then
                    lRs = lCt
                    Exit For
                End If
            End If
        Next lCt
    End If
    
    If (lRs > -1) Then
        lSp = lRs
        Find = lRs
        If bDisplay Then
            If m_bUseSpannedRows Then
                lSp = (m_lRowDepth(lRs) / m_lRowHeight) - 1
            End If
            RowEnsureVisible lSp
            If m_bFullRowSelect Then
                RowFocused(lRs) = True
            Else
                CellFocused(lRs, lColumn) = True
            End If
            GridRefresh False
        End If
    End If

On Error GoTo 0
Exit Function

Handler:

End Function



'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Function FindLongestString(ByVal lColumn As Long) As Long
'/* get longest cell string in column

Dim lRc As Long
Dim lUb As Long
Dim lLn As Long
Dim lTx As Long

    If (lColumn < m_lColumnCount) Then
        lUb = UBound(m_cGridItem)
        Do
            lTx = Len(m_cGridItem(lRc).Text(lColumn))
            If (lTx > lLn) Then
                lLn = lTx
            End If
            lRc = lRc + 1
        Loop Until (lRc > lUb)
        FindLongestString = lLn
    End If

End Function

Private Function FindShortestString(ByVal lColumn As Long) As Long
'/* get shortest cell string in column

Dim lRc As Long
Dim lUb As Long
Dim lLn As Long
Dim lTx As Long

    If (lColumn < m_lColumnCount) Then
        lUb = UBound(m_cGridItem)
        Do
            lTx = Len(m_cGridItem(lRc).Text(lColumn))
            If (lTx > 1) Then
                If (lTx < lLn) Then
                    lLn = lTx
                ElseIf (lLn = 0) Then
                    lLn = lTx
                End If
            End If
            lRc = lRc + 1
        Loop Until (lRc > lUb)
        FindShortestString = lLn
    End If

End Function

Private Sub FirstColumnWidth(ByVal lRow As Long)
'/* calculate first column width

Dim lIcon       As Long
Dim lColWidth   As Long
Dim sText       As String

    If m_bVirtualMode Then
        RaiseEvent eVHVirtualAccess(lRow, 0, sText, lIcon)
    Else
        lIcon = m_cGridItem(lRow).Icon(0)
    End If
    If (lIcon > -1) Then
        lColWidth = (m_lRowIconX + 6)
        m_bFirstRowReserved = True
    End If
    
    If m_bCheckBoxes Then
        lColWidth = lColWidth + 22
    End If
    If Not (lColWidth = 0) Then
        ColumnWidth(0) = lColWidth
    End If
    If m_bLockFirstColumn Then
        ColumnLock(0) = True
    End If

End Sub


'**********************************************************************
'*                              SORTING
'**********************************************************************

Private Function ArrayCheck(ByRef vArray As Variant) As Boolean
'/* validity test

On Error Resume Next

    '/* an array
    If Not IsArray(vArray) Then
        GoTo Handler
        '/* not dimensioned
    ElseIf IsError(UBound(vArray)) Then
        GoTo Handler
        '/* no members
    ElseIf (UBound(vArray) = -1) Then
        GoTo Handler
    End If
    ArrayCheck = True

Handler:
    On Error GoTo 0

End Function

Private Function BuildNumericSortArray(ByVal lColumn As Long) As Boolean
'/* create a copy of sort items

Dim lUb As Long
Dim lLb As Long
Dim lCt As Long

    If (lColumn < m_lColumnCount) Then
        Erase m_sSortArray
        Erase m_lSortArray
        lCt = 0
        If m_bFiltered Then
            lLb = LBound(m_lFilter)
            lUb = UBound(m_lFilter)
            ReDim m_lSortArray(lLb To lUb)
        Else
            lLb = LBound(m_cGridItem)
            lUb = UBound(m_cGridItem)
            ReDim m_lSortArray(lLb To lUb)
        End If
        '/* date
        If (m_eSortTag = ecsSortDate) Then
            If m_bFiltered Then
                Do
                    m_lSortArray(lCt) = GetTime(m_cGridItem(m_lFilter(lCt)).Text(lColumn))
                    lCt = lCt + 1
                Loop Until (lCt > lUb)
            Else
                Do
                    m_lSortArray(lCt) = GetTime(m_cGridItem(lCt).Text(lColumn))
                    lCt = lCt + 1
                Loop Until (lCt > lUb)
            End If
        '/* number
        ElseIf (IsNumeric(m_cGridItem(lCt).Text(lColumn))) Then
            If m_bFiltered Then
                Do
                    m_lSortArray(lCt) = CLng(m_cGridItem(m_lFilter(lCt)).Text(lColumn))
                    lCt = lCt + 1
                Loop Until (lCt > lUb)
            Else
                Do
                    m_lSortArray(lCt) = CLng(m_cGridItem(lCt).Text(lColumn))
                    lCt = lCt + 1
                Loop Until (lCt > lUb)
            End If
        '/* icon
        ElseIf (m_eSortTag = ecsSortIcon) Then
            If m_bFiltered Then
                Do
                    m_lSortArray(lCt) = m_cGridItem(m_lFilter(lCt)).Icon(lColumn)
                    lCt = lCt + 1
                Loop Until (lCt > lUb)
            Else
                Do
                    m_lSortArray(lCt) = m_cGridItem(lCt).Icon(lColumn)
                    lCt = lCt + 1
                Loop Until (lCt > lUb)
            End If
        '/* invalid
        Else
            GoTo Handler
        End If
        BuildNumericSortArray = True
    End If

On Error GoTo 0
Exit Function

Handler:
    BuildNumericSortArray = False
    On Error GoTo 0

End Function

Private Sub BuildStringSortArray(ByVal lColumn As Long)
'/* create a copy of sort items

Dim lUb As Long
Dim lLb As Long
Dim lCt As Long

    Erase m_sSortArray
    Erase m_lSortArray
    
    If m_bFiltered Then
        lLb = LBound(m_lFilter)
        lUb = UBound(m_lFilter)
        ReDim m_sSortArray(lLb To lUb)
        Do
            m_sSortArray(lCt) = m_cGridItem(m_lFilter(lCt)).Text(lColumn)
            If (LenB(m_sSortArray(lCt)) = 0) Then
                m_sSortArray(lCt) = "a"
            End If
            lCt = lCt + 1
        Loop Until (lCt > lUb)
    Else
        lLb = LBound(m_cGridItem)
        lUb = UBound(m_cGridItem)
        ReDim m_sSortArray(lLb To lUb)
        Do
            m_sSortArray(lCt) = m_cGridItem(lCt).Text(lColumn)
            If (LenB(m_sSortArray(lCt)) = 0) Then
                m_sSortArray(lCt) = "a"
            End If
            lCt = lCt + 1
        Loop Until (lCt > lUb)
    End If

End Sub

Private Sub ResetArray(ByRef cArray() As clsGridItem)
'/* reset array with new dimensions
'TODO

Dim lCt As Long
Dim lLb As Long
Dim lUb As Long
Dim lVl As Long

    If Not IsArray(cArray) Then Exit Sub
    lLb = LBound(cArray)
    lUb = UBound(cArray)

    If (lUb = -1) Or (lUb - lLb = 0) Then
        Erase cArray
        Exit Sub
    End If

    lVl = 0
    For lCt = lLb To lUb
        If Not (cArray(lCt) Is Nothing) Then
            Set cArray(lVl) = cArray(lCt)
            lVl = lVl + 1
        End If
    Next lCt
    ReDim Preserve cArray(lVl - 1)

End Sub

Private Sub ResizeArray(ByRef cArray() As clsGridItem, _
                        ByVal lPos As Long)

'/* redimension array
'TODO

Dim lCt As Long
Dim lLb As Long
Dim lUb As Long

    If Not IsArray(cArray) Then Exit Sub
    lLb = LBound(cArray)
    lUb = UBound(cArray)
    If (lUb = -1) Or (lUb - lLb = 0) Then
        Erase cArray
        Exit Sub
    End If
    '/* if invalid Pos
    If (lPos > lUb) Or (lPos = -1) Then
        lPos = lUb
    ElseIf (lPos < lLb) Then
        lPos = lLb
    ElseIf (lPos = lUb) Then
        ReDim Preserve cArray(lUb - 1)
        Exit Sub
    End If

    Set cArray(lPos) = Nothing
    For lCt = lPos + 1 To lUb
        Set cArray(lCt - 1) = cArray(lCt)
    Next lCt
    ReDim Preserve cArray(lUb - 1)

End Sub

Private Function GetTime(ByVal vDate As Variant) As Long
'/* get date from number

Dim lRet As Long

    If IsDate(vDate) Then
        lRet = Format$(vDate, "General Number")
    Else
        lRet = 0
    End If
    GetTime = lRet

End Function

Private Sub GridArrayResize(ByVal lRowCount As Long)
'/* resize grid items

Dim lLb As Long
Dim lUb As Long
Dim lCt As Long

    If ArrayCheck(m_cGridItem) Then
        lLb = LBound(m_cGridItem)
        lUb = UBound(m_cGridItem) + 1
    Else
        lLb = 0
        lUb = 0
    End If
    ReDim Preserve m_cGridItem(lLb To lRowCount)
    For lCt = lUb To lRowCount
        Set m_cGridItem(lCt) = New clsGridItem
        m_cGridItem(lCt).Init m_lColumnCount - 1
    Next lCt
    m_lRowCount = UBound(m_cGridItem) + 1

End Sub

Public Function ItemsSort(ByVal lColumn As Long, _
                          ByVal bDescending As Boolean) As Boolean

'*/ sort items in the list

    If Not m_bVirtualMode Then
        If bDescending Then
            SortControl False, lColumn
        Else
            SortControl True, lColumn
        End If
        SetRowCount Count
    End If

End Function

Private Function MoveArrayItem(ByVal lIntPos As Long, _
                               ByVal lDstPos As Long) As Boolean

'/* shift an item in array

Dim lCt     As Long
Dim lInPtr  As Long
Dim lDtPtr  As Long
Dim lTpPtr  As Long
Dim cLTemp  As clsGridItem

    If (lIntPos > lDstPos) Then
        Set cLTemp = m_cGridItem(lIntPos)
        lTpPtr = VarPtr(cLTemp)
        For lCt = lIntPos To (lDstPos + 1) Step -1
            lInPtr = VarPtr(m_cGridItem(lCt - 1))
            lDtPtr = VarPtr(m_cGridItem(lCt))
            CopyMemBv lDtPtr, lInPtr, 4&
        Next lCt
        CopyMemBv lDtPtr, lTpPtr, 4&
        CopyMemBr ByVal lTpPtr, 0&, 4&
        Set cLTemp = Nothing
    ElseIf (lIntPos < lDstPos) Then
        Set cLTemp = m_cGridItem(lIntPos)
        lTpPtr = VarPtr(cLTemp)
        For lCt = lIntPos To (lDstPos - 1)
            lInPtr = VarPtr(m_cGridItem(lCt))
            lDtPtr = VarPtr(m_cGridItem(lCt + 1))
            CopyMemBv lInPtr, lDtPtr, 4&
        Next lCt
        CopyMemBv lDtPtr, lTpPtr, 4&
        CopyMemBr ByVal lTpPtr, 0&, 4&
        Set cLTemp = Nothing
    End If
    m_bSorted = False
    If m_bUseSpannedRows Then
        RowArrayResize
    End If

End Function

Private Sub QSINumericSort(ByRef aLTmp() As Long, _
                           ByVal bDesc As Boolean, _
                           Optional ByVal lLb As Long = -1, _
                           Optional ByVal lUb As Long = -1)

'/* sort longs from temp array, and shuffle grid items in parallel
'/* adaptation of Rohan Edwards' StrSwap4 quicksort:
'/* http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=63800&lngWId=1

Dim lLow    As Long
Dim lHigh   As Long
Dim lCount  As Long
Dim lSPtr   As Long
Dim lAPtr   As Long
Dim lTPtr   As Long
Dim lCAPtr  As Long
Dim lItem   As Long
Dim cItem   As clsGridItem

    If (lLb = -1) Then
        lLb = LBound(aLTmp)
    End If
    If (lUb = -1) Then
        lUb = UBound(aLTmp)
    End If
    
    Set cItem = New clsGridItem
    lHigh = (((lUb - lLb) \ 8&) + 32&)
    ReDim lbs(1& To lHigh) As Long
    ReDim ubs(1& To lHigh) As Long
    
    lSPtr = VarPtr(lItem)
    lAPtr = VarPtr(aLTmp(lLb)) - (lLb * 4&)
    '*/ temp pointer
    lTPtr = VarPtr(cItem)
    '*/ array pointer
    lCAPtr = VarPtr(m_cGridItem(lLb)) - (lLb * 4&)
    
    Do
        lHigh = ((lUb - lLb) \ 2&) + lLb
        CopyMemBv lSPtr, lAPtr + (lHigh * 4&), 4&
        CopyMemBv lAPtr + (lHigh * 4&), lAPtr + (lUb * 4&), 4&
        '*/ copy array lItem to temp
        CopyMemBv lTPtr, lCAPtr + (lHigh * 4&), 4&
        '*/ swap array items
        CopyMemBv lCAPtr + (lHigh * 4&), lCAPtr + (lUb * 4&), 4&
        lLow = lLb
        lHigh = lUb
        Do While (lHigh > lLow)
            If (lItem > aLTmp(lLow)) = bDesc Then
                CopyMemBv lAPtr + (lHigh * 4&), lAPtr + (lLow * 4&), 4&
                '*/ swap array items
                CopyMemBv lCAPtr + (lHigh * 4&), lCAPtr + (lLow * 4&), 4&
                lHigh = lHigh - 1&
                Do Until (lHigh = lLow)
                    If (aLTmp(lHigh) > lItem) = bDesc Then
                        CopyMemBv lAPtr + (lLow * 4&), lAPtr + (lHigh * 4&), 4&
                        '*/ swap grid array items
                        CopyMemBv lCAPtr + (lLow * 4&), lCAPtr + (lHigh * 4&), 4&
                        Exit Do
                    End If
                    lHigh = (lHigh - 1&)
                Loop
                If (lHigh = lLow) Then Exit Do
            End If
            lLow = (lLow + 1&)
        Loop
        
        CopyMemBv lAPtr + (lHigh * 4&), lSPtr, 4&
        '*/ copy temp to array
        CopyMemBv lCAPtr + (lHigh * 4&), lTPtr, 4&
        
        If (lLb < lLow - 1&) Then
            If (lUb > lLow + 1&) Then
                lCount = lCount + 1&
                lbs(lCount) = lLow + 1&
                ubs(lCount) = lUb
            End If
            lUb = (lLow - 1&)
        ElseIf (lUb > lLow + 1&) Then
            lLb = (lLow + 1&)
        Else
            If lCount = 0& Then Exit Do
            lLb = lbs(lCount)
            lUb = ubs(lCount)
            lCount = lCount - 1&
        End If
    Loop
    CopyMemBr ByVal lSPtr, 0&, 4&
    '*/ cleanup
    CopyMemBr ByVal lTPtr, 0&, 4&
    Set cItem = Nothing

End Sub

Private Sub QSIStringSort(ByRef aSTmp() As String, _
                          ByVal eDirection As eSDSortDirection, _
                          ByVal lCase As VbCompareMethod, _
                          Optional ByVal lLb As Long = -1, _
                          Optional ByVal lUb As Long = -1)

'/* sort strings from temp array, and shuffle grid items in parallel
'/* adaptation of Rohan Edwards' StrSwap4 quicksort:
'/* http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=63800&lngWId=1

Dim lLow    As Long
Dim lHigh   As Long
Dim lCount  As Long
Dim lSPtr   As Long
Dim lAPtr   As Long
Dim lTPtr   As Long
Dim lCAPtr  As Long
Dim sItem   As String
Dim cItem   As clsGridItem

    If (lLb = -1) Then
        lLb = LBound(aSTmp)
    End If
    If (lUb = -1) Then
        lUb = UBound(aSTmp)
    End If
    
    Set cItem = New clsGridItem
    lHigh = ((lUb - lLb) \ 8&) + 32&
    ReDim lbs(1& To lHigh) As Long
    ReDim ubs(1& To lHigh) As Long
    lSPtr = VarPtr(sItem)
    lAPtr = VarPtr(aSTmp(lLb)) - (lLb * 4&)
    '*/ temp pointer
    lTPtr = VarPtr(cItem)
    '*/ array pointer
    lCAPtr = VarPtr(m_cGridItem(lLb)) - (lLb * 4&)

    Do
        lHigh = ((lUb - lLb) \ 2&) + lLb
        CopyMemBv lSPtr, lAPtr + (lHigh * 4&), 4&
        CopyMemBv lAPtr + (lHigh * 4&), lAPtr + (lUb * 4&), 4&
        '*/ copy array sItem to temp
        CopyMemBv lTPtr, lCAPtr + (lHigh * 4&), 4&
        '*/ swap array items
        CopyMemBv lCAPtr + (lHigh * 4&), lCAPtr + (lUb * 4&), 4&
        lLow = lLb
        lHigh = lUb
        Do While (lHigh > lLow)
            If Not (StrComp(sItem, aSTmp(lLow), lCase) = eDirection) Then
                CopyMemBv lAPtr + (lHigh * 4&), lAPtr + (lLow * 4&), 4&
                '*/ swap grid array items
                CopyMemBv lCAPtr + (lHigh * 4&), lCAPtr + (lLow * 4&), 4&
                lHigh = lHigh - 1&
                Do Until (lHigh = lLow)
                    If Not (StrComp(aSTmp(lHigh), sItem, lCase) = eDirection) Then
                        CopyMemBv lAPtr + (lLow * 4&), lAPtr + (lHigh * 4&), 4&
                        '*/ swap grid array items
                        CopyMemBv lCAPtr + (lLow * 4&), lCAPtr + (lHigh * 4&), 4&
                        Exit Do
                    End If
                    lHigh = (lHigh - 1&)
                Loop
                If (lHigh = lLow) Then Exit Do
            End If
            lLow = (lLow + 1&)
        Loop
        
        CopyMemBv lAPtr + (lHigh * 4&), lSPtr, 4&
        '*/ copy temp to array
        CopyMemBv lCAPtr + (lHigh * 4&), lTPtr, 4&

        If (lLb < lLow - 1&) Then
            If (lUb > lLow + 1&) Then
                lCount = lCount + 1&
                lbs(lCount) = lLow + 1&
                ubs(lCount) = lUb
            End If
            lUb = (lLow - 1&)
        ElseIf (lUb > lLow + 1&) Then
            lLb = (lLow + 1&)
        Else
            If lCount = 0& Then Exit Do
            lLb = lbs(lCount)
            lUb = ubs(lCount)
            lCount = lCount - 1&
        End If
    Loop
    CopyMemBr ByVal lSPtr, 0&, 4&
    '*/ cleanup
    CopyMemBr ByVal lTPtr, 0&, 4&
    Set cItem = Nothing

End Sub

Private Sub SortControl(ByVal bAscending As Boolean, _
                        ByVal lColumn As Long, _
                        Optional ByVal lLb As Long = -1, _
                        Optional ByVal lUb As Long = -1)

'/* sorting hub

    '/* no sorting
    If (m_eSortType = -1) Then
        Exit Sub
    ElseIf (m_lRowCount = 0) Then
        Exit Sub
    End If
    
    m_eSortTag = ColumnTag(lColumn)
    '/* auto determine sort type
    If m_eSortTag = ecsSortAuto Then
        m_eSortTag = SortSample(lColumn)
    End If
    If m_bFiltered Then
        FilterReset
        m_bFiltered = False
    End If
    '/* build temp array
    Select Case m_eSortTag
        '/* string sort
    Case ecsSortDefault
        BuildStringSortArray lColumn
        If (UBound(m_sSortArray) < 2) Then
            GoTo Handler
        End If
        '/* array less then min dimensions
        If ArrayCheck(m_sSortArray) Then
            If bAscending Then
                QSIStringSort m_sSortArray, esdAscending, m_eSortType, lLb, lUb
            Else
                QSIStringSort m_sSortArray, esdDescending, m_eSortType, lLb, lUb
            End If
        End If
    '/* numeric and date
    Case ecsSortDate, ecsSortNumeric, ecsSortIcon
        If BuildNumericSortArray(lColumn) Then
            If (UBound(m_lSortArray) < 2) Then
                GoTo Handler
            End If
            If ArrayCheck(m_lSortArray) Then
                If bAscending Then
                    QSINumericSort m_lSortArray, False, lLb, lUb
                Else
                    QSINumericSort m_lSortArray, True, lLb, lUb
                End If
            End If
        End If
    End Select
    
    If m_bUseSpannedRows Then
        RowArrayResize
    End If
    If m_bHasSubCells Then
        SubCellSortRows
    End If
    GridRefresh False
    '/* success
    m_bSorted = True

Handler:

End Sub

Private Function SortSample(ByVal lColumn As Long) As ECSColumnSortTags
'/* auto determine sort type

Dim vItem As Variant

    vItem = m_cGridItem(RowTopIndex).Text(lColumn)
    If (LenB(vItem) = 0) Then
        SortSample = ecsSortNone
    ElseIf IsDate(vItem) Then
        SortSample = ecsSortDate
    ElseIf IsNumeric(vItem) Then
        SortSample = ecsSortNumeric
    ElseIf (LenB(vItem) > 0) Then
        SortSample = ecsSortDefault
    Else
        SortSample = ecsSortNone
    End If

End Function

Public Sub SubSort(ByVal bAscending As Boolean, _
                   ByVal lColumn As Long, _
                   ByVal lStartPos As Long, _
                   ByVal lEndPos As Long)

'/* sotrt items within range

    If Not m_bVirtualMode Then
        If (lColumn < m_lColumnCount) Then
            If (lEndPos < m_lRowCount) Then
                SortControl bAscending, lColumn, lStartPos, lEndPos
            End If
        End If
    End If

End Sub


'**********************************************************************
'*                              EDIT CONTROL
'**********************************************************************
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/commctls/editcontrols/abouteditcontrols.asp
Private Sub AdvancedEditActivate()

Dim lIcon   As Long
Dim sText   As String
Dim tRect   As RECT
Dim tRClt   As RECT

    If m_bItemActive Then
        EditUpdateLabel
    End If
    '/* selected cell and row
    m_lEditItem = RowHitTest
    m_lEditSubItem = CellHitTest
    If (m_lEditItem > -1) Then
        If (m_lEditSubItem > -1) Then
            If Not m_bEditorLoaded Then
                m_bEditorLoaded = True
                If m_bHasSubCells Then
                    GridRefresh False
                End If
                CellCalcRect m_lEditItem, m_lEditSubItem, tRect
                GetWindowRect m_lHGHwnd, tRClt
                Set m_cEditor = New clsAdvancedEdit
                If m_bVirtualMode Then
                    RaiseEvent eVHAdvancedEditRequestText(m_lEditItem, m_lEditSubItem, lIcon, sText)
                Else
                    With m_cGridItem(m_lEditItem)
                        lIcon = .Icon(m_lEditSubItem)
                        sText = .Text(m_lEditSubItem)
                    End With
                End If
                With tRClt
                    .left = .left + tRect.left
                    .top = .top + tRect.top
                    m_cEditor.FontRightLeading = m_bFontRightLeading
                    m_cEditor.UseUnicode = m_bUseUnicode
                    m_cEditor.CreateEditBox m_lHGHwnd, .left, .top, m_eEditorThemeStyle, m_lAdvancedEditThemeColor, m_lAdvancedEditOffsetColor, sText, m_lImlRowHndl, lIcon, eiaIcon
                End With
                m_bEditorLoaded = True
            End If
        End If
    End If

End Sub

Private Function AdvancedEditorRequest() As Boolean

    If CBool(GetAsyncKeyState(VK_CONTROL)) Then
        AdvancedEditorRequest = True
    End If
    
End Function

Private Sub EditBoxActivate()
'/* launch edit window

Dim sText As String
Dim tRect As RECT

    If m_bItemActive Then
        EditUpdateLabel
    End If
    '/* selected cell and row
    m_lEditItem = RowHitTest
    m_lEditSubItem = CellHitTest
    If (m_lEditItem > -1) Then
        If (m_lEditSubItem > -1) Then
            If m_bVirtualMode Then
                RaiseEvent eVHEditRequestText(m_lEditItem, m_lEditSubItem, sText)
            Else
                sText = m_cGridItem(m_lEditItem).Text(m_lEditSubItem)
            End If
            If (LenB(sText) > 0) Then
                StopTipTimer
                m_bItemActive = True
                If m_bHasSubCells Then
                    SubCellsHideAll
                End If
                '/* refresh cell
                CellErase m_lEditItem, m_lEditSubItem, False
                '/* launch edit window
                CellCalcRect m_lEditItem, m_lEditSubItem, tRect
                EditControlCreate m_lEditItem, m_lEditSubItem, sText
                EditShow True
            End If
        End If
    End If

End Sub

Public Property Get EditControlType() As ECTEditControlType
Attribute EditControlType.VB_MemberFlags = "40"
'/* [get] edit control type
    EditControlType = m_eEditControlType
End Property

Public Property Let EditControlType(ByVal PropVal As ECTEditControlType)
'/* [let] edit control type
    m_eEditControlType = PropVal
End Property

Public Sub EditControlAddData(ByVal sData As String, _
                              Optional ByVal lIconHnd As Long, _
                              Optional ByVal lIconIdx As Long)

'/* add extra data to edit control

    If Not (m_cEditBox Is Nothing) Then
        Select Case m_eEditControlType
        Case ectTextBox
            m_cEditBox.Text = sData
        Case ectCombo, ectListbox
            m_cEditBox.AddItem sData
        Case ectImageCombo, ectImageListbox
            With m_cEditBox
                If (lIconHnd > 0) Then
                    .ImlListBoxAddIcon lIconHnd
                End If
                .AddItem sData, lIconIdx
            End With
        End Select
    End If
    
End Sub

Private Sub EditControlCreate(ByVal lRow As Long, _
                              ByVal lCell As Long, _
                              ByVal sText As String)
'/* create the edit control

Dim lCenter As Long
Dim tCntl   As RECT

    '/* destroy old instance
    If Not (m_cEditBox Is Nothing) Then
        EditControlDestroy
    End If
    Set m_cEditBox = New clsODControl
    '/* test client pos
    With m_tREditCrd
        If (.top < (m_lHeaderOffset)) Then
            .top = m_lHeaderOffset
        End If
    End With
    '/* create the control
    With m_cEditBox
        Select Case m_eEditControlType
        Case ectCombo, ectImageCombo
            .Name = "cbGridEdit"
            .BorderStyle ecbsThin
            .ThemeStyle = m_eHeaderSkinStyle
            If (m_oThemeColor > -1) Then
                .ThemeColor = m_oThemeColor
            End If
            .UseUnicode = m_bUseUnicode
            .HFont = m_lFont
            CellCalcRect lRow, lCell, tCntl
            With tCntl
                lCenter = .top + (((.Bottom - .top) - 25) / 2)
            End With
            If (m_eEditControlType = ectImageCombo) Then
                With m_tREditCrd
                    m_cEditBox.Create m_lHGHwnd, .left, lCenter, (.Right - .left) - 4, 160, ecsImageCombo
                End With
            Else
                With m_tREditCrd
                    m_cEditBox.Create m_lHGHwnd, .left, lCenter, (.Right - .left) - 4, 250, ecsComboDropDown
                End With
            End If
            .Visible = False
            .Text = m_cGridItem(m_lEditItem).Text(m_lEditSubItem)
            m_cGridItem(m_lEditItem).Text(m_lEditSubItem) = ""
        Case ectListbox, ectImageListbox
            .Name = "lstGridEdit"
            .BackColor = m_oBackColor
            .ForeColor = m_oForeColor
            .BorderStyle ecbsThin
            .UseUnicode = m_bUseUnicode
            .HFont = m_lFont
            If (m_eEditControlType = ectImageListbox) Then
                With m_tREditCrd
                    m_cEditBox.Create m_lHGHwnd, .left, .top, (.Right - .left), (.Bottom - .top), ecsImageListBox
                End With
            Else
                With m_tREditCrd
                    m_cEditBox.Create m_lHGHwnd, .left, .top, (.Right - .left), (.Bottom - .top), ecsListBox
                End With
            End If
            .Visible = False
            .BorderStyle ecbsThin
        Case ectTextBox
            .Name = "txtGridEdit"
            .UseUnicode = m_bUseUnicode
            .NoScrollbar = True
            .BackColor = m_oBackColor
            .ForeColor = m_oForeColor
            .HFont = m_lFont
            .BorderStyle ecbsThin
            With m_tREditCrd
                m_cEditBox.Create m_lHGHwnd, .left, .top, (.Right - .left), (.Bottom - .top), ecsTextBox
            End With
            .Visible = True
            .Text = sText
        End Select
        m_lEditHwnd = .hwnd
    End With
    '/* flag the event
    RaiseEvent eVHEditorLoaded(m_eEditControlType, lRow, lCell)

End Sub

Private Sub EditControlDestroy()
'/* destroy editor

    Set m_cEditBox = Nothing
    m_lEditHwnd = 0
    
End Sub

Private Function EditGetText() As String
'/* get edit box text

Dim lLen  As Long
Dim sText As String

    If Not (m_lEditHwnd = 0) Then
        Select Case m_eEditControlType
        Case ectTextBox
            If m_bIsNt Then
                lLen = GetWindowTextLengthW(m_lEditHwnd)
            Else
                lLen = GetWindowTextLengthA(m_lEditHwnd)
            End If
            If (lLen > 0) Then
                lLen = lLen + 1
                sText = String$(lLen, Chr$(0))
                If m_bIsNt Then
                    GetWindowTextW m_lEditHwnd, StrPtr(sText), lLen
                Else
                    GetWindowTextA m_lEditHwnd, sText, lLen
                End If
            lLen = InStr(1, sText, vbNullChar)
                If (lLen > 0) Then
                    sText = left$(sText, (lLen - 1))
                End If
                EditGetText = sText
            Else
                EditGetText = ""
            End If
        Case ectCombo, ectImageCombo
            With m_cEditBox
                If (.ListIndex = -1) Then
                    EditGetText = .Text
                Else
                    EditGetText = .ListText(.ListIndex)
                End If
            End With
        Case ectListbox, ectImageListbox
            With m_cEditBox
                If Not (.ListIndex = -1) Then
                    EditGetText = .ListText(.ListIndex)
                End If
            End With
        End Select
    End If

End Function

Public Sub EditLimitLength(ByVal lLength As Long)
'/* limit edit box text length

    If (m_eEditControlType = ectTextBox) Then
        If Not (m_lEditHwnd = 0) Then
            SendMessageLongA m_lEditHwnd, EM_LIMITTEXT, lLength, 0&
        End If
    End If

End Sub

Public Sub EditLowerCase()
'/* set chars in edit box to lowercase

Dim lStyle   As Long
    
    If (m_eEditControlType = ectTextBox) Then
        If Not (m_lEditHwnd = 0) Then
            If m_bIsNt Then
                lStyle = GetWindowLongW(m_lEditHwnd, GWL_STYLE)
            Else
                lStyle = GetWindowLongA(m_lEditHwnd, GWL_STYLE)
            End If
            If m_bIsNt Then
                SetWindowLongW m_lEditHwnd, GWL_STYLE, lStyle Or ES_LOWERCASE
            Else
                SetWindowLongA m_lEditHwnd, GWL_STYLE, lStyle Or ES_LOWERCASE
            End If
        End If
    End If

End Sub

Private Sub EditShow(ByVal bVisible As Boolean)
'/* show edit box

    If Not (m_lEditHwnd = 0) Then
        If bVisible Then
            ShowWindow m_lEditHwnd, SW_NORMAL
        Else
            ShowWindow m_lEditHwnd, SW_HIDE
        End If
    End If

End Sub

Private Sub EditUpdateLabel()
'/* update listview cell

Dim sText As String

    If Not (IsWindowVisible(m_lEditHwnd) = 0) Then
        If m_bItemActive Then
            '/* get text
            sText = EditGetText
            RaiseEvent eVHEditChange(m_lEditItem, m_lEditSubItem, sText)
            If Not m_bVirtualMode Then
                '/* write to griditem array
                Select Case m_eEditControlType
                Case ectListbox, ectImageListbox
                    If Not (LenB(sText) = 0) Then
                        m_cGridItem(m_lEditItem).Text(m_lEditSubItem) = sText
                    End If
                Case Else
                    m_cGridItem(m_lEditItem).Text(m_lEditSubItem) = sText
                End Select
            End If
            '/* hide and clear editbox
            EditShow False
            m_lEditSubItem = -1
            m_lEditItem = -1
            SetFocus m_lHGHwnd
            CellRefresh RowInFocus, CellInFocus
            EditControlDestroy
            m_bItemActive = False
        End If
    End If

End Sub

Public Sub EditUpperCase()
'/* edit box chars all uppercase

Dim lStyle   As Long

    If Not (m_lEditHwnd = 0) Then
        If (m_eEditControlType = ectTextBox) Then
            If m_bUseUnicode Then
                lStyle = GetWindowLongW(m_lEditHwnd, GWL_STYLE)
            Else
                lStyle = GetWindowLongA(m_lEditHwnd, GWL_STYLE)
            End If
            If m_bIsNt Then
                SetWindowLongW m_lEditHwnd, GWL_STYLE, lStyle Or ES_UPPERCASE
            Else
                SetWindowLongA m_lEditHwnd, GWL_STYLE, lStyle Or ES_UPPERCASE
            End If
        End If
    End If

End Sub


'**********************************************************************
'*                              SUBCLASSING
'**********************************************************************

Private Sub GridAttatch()
'/* attatch messages

    If Not (m_lHGHwnd = 0) Then
        With m_cHGridSubclass
            If Not (m_lParentHwnd = 0) Then
                .Subclass m_lParentHwnd, Me
                .AddMessage m_lParentHwnd, WM_DRAWITEM, MSG_BEFORE
                .AddMessage m_lParentHwnd, WM_MEASUREITEM, MSG_BEFORE
                .AddMessage m_lParentHwnd, WM_NOTIFY, MSG_BEFORE
                .AddMessage m_lParentHwnd, WM_SETFOCUS, MSG_BEFORE
                .AddMessage m_lParentHwnd, WM_SIZE, MSG_BEFORE
                .AddMessage m_lParentHwnd, WM_DISPLAYCHANGE, MSG_AFTER
                .AddMessage m_lParentHwnd, WM_SETTINGCHANGE, MSG_AFTER
                .AddMessage m_lParentHwnd, WM_MOUSEMOVE, MSG_BEFORE
                .AddMessage m_lParentHwnd, WM_TIMER, MSG_BEFORE
                .AddMessage m_lParentHwnd, WM_LBUTTONUP, MSG_BEFORE
            End If
            .Subclass m_lHGHwnd, Me
            .AddMessage m_lHGHwnd, WM_CTLCOLOREDIT, MSG_BEFORE
            .AddMessage m_lHGHwnd, WM_KEYDOWN, MSG_BEFORE
            .AddMessage m_lHGHwnd, WM_ERASEBKGND, MSG_BEFORE
            .AddMessage m_lHGHwnd, WM_LBUTTONDOWN, MSG_BEFORE
            .AddMessage m_lHGHwnd, WM_LBUTTONUP, MSG_BEFORE
            .AddMessage m_lHGHwnd, WM_MOUSEMOVE, MSG_BEFORE
            .AddMessage m_lHGHwnd, WM_NCLBUTTONDOWN, MSG_BEFORE
            .AddMessage m_lHGHwnd, WM_SETCURSOR, MSG_BEFORE
            .AddMessage m_lHGHwnd, WM_MOUSEWHEEL, MSG_BEFORE
            .AddMessage m_lHGHwnd, WM_CHAR, MSG_BEFORE
            If Not (m_lHdrHwnd = 0) Then
                .Subclass m_lHdrHwnd, Me
                .AddMessage m_lHdrHwnd, HDM_LAYOUT, MSG_BEFORE
                .AddMessage m_lHdrHwnd, WM_LBUTTONUP, MSG_BEFORE
                .AddMessage m_lHdrHwnd, WM_MOUSEMOVE, MSG_BEFORE
            End If
        End With
    End If

End Sub

Private Sub GridDetatch()
'/* detatch messages

    If Not (m_lHGHwnd = 0) Then
        With m_cHGridSubclass
            If Not (m_lParentHwnd = 0) Then
                .DeleteMessage m_lParentHwnd, WM_DRAWITEM, MSG_BEFORE
                .DeleteMessage m_lParentHwnd, WM_MEASUREITEM, MSG_BEFORE
                .DeleteMessage m_lParentHwnd, WM_NOTIFY, MSG_BEFORE
                .DeleteMessage m_lParentHwnd, WM_SETFOCUS, MSG_BEFORE
                .DeleteMessage m_lParentHwnd, WM_SIZE, MSG_BEFORE
                .DeleteMessage m_lParentHwnd, WM_DISPLAYCHANGE, MSG_AFTER
                .DeleteMessage m_lParentHwnd, WM_SETTINGCHANGE, MSG_AFTER
                .DeleteMessage m_lParentHwnd, WM_MOUSEMOVE, MSG_BEFORE
                .DeleteMessage m_lParentHwnd, WM_TIMER, MSG_BEFORE
                .DeleteMessage m_lParentHwnd, WM_LBUTTONUP, MSG_BEFORE
                TreeRemoveSizer
                .UnSubclass m_lParentHwnd
            End If
            If Not (m_lHGHwnd = 0) Then
                .DeleteMessage m_lHGHwnd, WM_CTLCOLOREDIT, MSG_BEFORE
                .DeleteMessage m_lHGHwnd, WM_KEYDOWN, MSG_BEFORE
                .DeleteMessage m_lHGHwnd, WM_ERASEBKGND, MSG_BEFORE
                .DeleteMessage m_lHGHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                .DeleteMessage m_lHGHwnd, WM_LBUTTONUP, MSG_BEFORE
                .DeleteMessage m_lHGHwnd, WM_MOUSEMOVE, MSG_BEFORE
                .DeleteMessage m_lHGHwnd, WM_NCLBUTTONDOWN, MSG_BEFORE
                .DeleteMessage m_lHGHwnd, WM_SETCURSOR, MSG_BEFORE
                .DeleteMessage m_lHGHwnd, WM_MOUSEWHEEL, MSG_BEFORE
                .DeleteMessage m_lHGHwnd, WM_CHAR, MSG_BEFORE
                .UnSubclass m_lHGHwnd
            End If
            If Not (m_lHdrHwnd = 0) Then
                .DeleteMessage m_lHdrHwnd, HDM_LAYOUT, MSG_BEFORE
                .DeleteMessage m_lHdrHwnd, WM_LBUTTONUP, MSG_BEFORE
                .DeleteMessage m_lHdrHwnd, WM_MOUSEMOVE, MSG_BEFORE
                .UnSubclass m_lHdrHwnd
            End If
        End With
    End If

End Sub

Private Function StartTipTimer() As Boolean
'/* start display timer

    If Not m_bTipTimerActive Then
        SetTimer m_lParentHwnd, 1&, 100&, 0&
        m_bTipTimerActive = True
        m_bTipTracking = True
    End If

End Function

Private Function StopTipTimer() As Boolean
'/* stop display timer

    If m_bTipTimerActive Then
        KillTimer m_lParentHwnd, 1&
        m_cCellTips.DestroyToolTip
        m_bTipTimerActive = False
        m_bTipTracking = False
        m_bShowing = False
        m_lTipTimer = 0
    End If

End Function

Private Sub HeaderDragStop()

    If m_bDragDrop Then
        If m_bColumnDragging Then
            ColumnDragging = False
            If Not (m_lDraggedColumn = -1) Then
                ColumnIconStore m_lDraggedColumn, True
            End If
            DragStopTimer
            HotDividerReset
            m_lDraggedColumn = -1
            RaiseEvent eVHColumnDragComplete
        End If
    End If

End Sub

Private Sub GXISubclass_WndProc(ByVal bBefore As Boolean, _
                                ByRef bHandled As Boolean, _
                                ByRef lReturn As Long, _
                                ByVal lHwnd As Long, _
                                ByVal uMsg As eMsg, _
                                ByVal wParam As Long, _
                                ByVal lParam As Long, _
                                ByRef lParamUser As Long)
'/* subclassing callback

Dim bDesc       As Boolean
Dim lZHnd       As Long
Dim lOrigin     As Long
Dim lCode       As Long
Dim lColumn     As Long
Dim tDIstc      As DRAWITEMSTRUCT
Dim tHDL        As HDLAYOUT
Dim tMsItm      As MEASUREITEMSTRUCT
Dim tNmhdr      As NMHDR
Dim tHDN        As NMHEADER
Dim tNmList     As NMLISTVIEW
Dim uNMTV       As NMTREEVIEW
Dim tPaint      As PAINTSTRUCT
Dim tRect       As RECT
Dim tWPos       As WINDOWPOS

    Select Case uMsg
    '/* draw notification
    Case WM_DRAWITEM
        '/* owner drawn grid
        If m_bDraw Then
            CopyMemory tDIstc, ByVal lParam, LenB(tDIstc)
            If Not m_bTransitionMask Then
                With tDIstc
                    DrawGrid .hdc, .itemID
                End With
            Else
                DrawTransitionMask tDIstc.hdc
            End If
            bHandled = True
        End If
            
    Case WM_PAINT
        If Not m_bPainting Then
            m_bPainting = True
            BeginPaint lHwnd, tPaint
            DrawSizer UserControl.hdc
            EndPaint lHwnd, tPaint
            m_bPainting = False
            bHandled = True
        Else
            lReturn = m_cHGridSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        End If
    
    '/* controlled erase
    Case WM_ERASEBKGND
        If m_bHasTreeView Then
            DrawSizer UserControl.hdc
            lReturn = 1
        ElseIf m_bHasInitialized Then
            lReturn = 1
        End If
        
    '/* list click
    Case WM_LBUTTONDOWN
        If m_bCellTips Then
            StopTipTimer
        End If
        HeaderHitState = HdrOffColumn
        '/* transition reset
        m_bTransitionMask = False
        '/* grid in focus
        m_bGridFocused = True
        '/* close subitem edit
        If m_bItemActive Then
            EditUpdateLabel
        End If
        If m_bCheckBoxes Then
            CheckBoxHitTest
        End If
        '/* using column focus
        If m_bColumnFocus Then
            GridRefresh False
        End If
        '/* redraw selected
        CellRedraw
    
    '/* header height
    Case HDM_LAYOUT
        If Not m_bHeaderHide Then
            '/! hey m$, where's the f***ing documentation?
            CopyMemory tHDL, ByVal lParam, Len(tHDL)
            CopyMemory tWPos, ByVal tHDL.lpwpos, Len(tWPos)
            CopyMemory tRect, ByVal tHDL.lprc, Len(tRect)
            '/* setwinpos struct
            With tWPos
                .hwnd = m_lHdrHwnd
                .hWndInsertAfter = 0&
                .cx = (tRect.Right - tRect.left)
                .cy = m_lHeaderHeight
                .x = tRect.left
                .y = 0&
                .flags = SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_NOMOVE
            End With
            '/* adjust client rect
            tRect.top = m_lHeaderHeight + 2
            '/* copy and forward
            CopyMemory ByVal tHDL.lprc, tRect, Len(tRect)
            CopyMemory ByVal tHDL.lpwpos, tWPos, Len(tWPos)
            CopyMemory ByVal lParam, tHDL, Len(tHDL)
           ' lReturn = lParam
            bHandled = True
        End If
    
    '/* release transition mask
    Case WM_LBUTTONUP
        If (lHwnd = m_lParentHwnd) Then
            If m_bHasTreeView Then
                If TreeViewDividerHitTest Then
                    m_cTreeView.Refresh True
                    DrawSizer UserControl.hdc
                End If
            End If
        Else
            ColumnSizingVertical = False
            HeaderDragStop
            If m_bTransitionMask Then
                DestroyTransitionMask
            End If
        End If
    
    Case WM_KEYDOWN
        Select Case wParam
        '/* tab
        Case VK_TAB
            If m_bItemActive Then
                EditUpdateLabel
            End If
            '/* next in z-order
            lZHnd = GetWindow(m_lParentHwnd, GW_HWNDNEXT)
            '/* focus
            If Not (lZHnd = 0) Then
                SetFocus lZHnd
                RePaint
                bHandled = True
            Else
                '/* return to parent
                lZHnd = GetParent(m_lParentHwnd)
                If Not (lZHnd = 0) Then
                    SetFocus lZHnd
                    RePaint
                    bHandled = True
                End If
            End If
        '/* esc key
        Case VK_ESCAPE
            If m_bItemActive Then
                EditUpdateLabel
                bHandled = True
            Else
                '/* return to parent
                lZHnd = GetParent(m_lParentHwnd)
                If Not (lZHnd = 0) Then
                    SetFocus lZHnd
                    bHandled = True
                End If
            End If
        '/* direction keys
        Case VK_LEFT
            If m_bItemActive Then
                EditUpdateLabel
            End If
            If Not m_bVirtualMode Then
                If Not m_bFullRowSelect Then
                    CellKeyNavigate EKNLeft
                End If
                bHandled = True
            End If
        Case VK_RIGHT
            If m_bItemActive Then
                EditUpdateLabel
            End If
            If Not m_bVirtualMode Then
                If Not m_bFullRowSelect Then
                    CellKeyNavigate EKNRight
                End If
                bHandled = True
            End If
        Case VK_UP
            If m_bItemActive Then
                EditUpdateLabel
            End If
            If Not m_bVirtualMode Then
                CellKeyNavigate EKNUp
                bHandled = True
            End If
        Case VK_DOWN
            If m_bItemActive Then
                EditUpdateLabel
            End If
            If Not m_bVirtualMode Then
                CellKeyNavigate EKNDown
                bHandled = True
            End If
        '/* enter key
        Case VK_ENTER
            If m_bItemActive Then
                If Not (m_eEditControlType = ectTextBox) Then
                    EditUpdateLabel
                End If
            End If
        '/* space bar
        Case VK_SPACE
            If m_bItemActive Then
                EditUpdateLabel
            End If
            If m_bCheckBoxes Then
                CheckToggle RowInFocus
                CellRefresh RowInFocus
            End If
        '/* 'a' char
        Case VK_UCASEA
            '/* advanced edit
            If m_bAdvancedEdit Then
                If Not m_cGridItem(RowInFocus).RowNoEdit Then
                    If AdvancedEditorRequest Then
                        RaiseEvent eVHAdvancedEditRequest(RowInFocus, CellInFocus)
                        AdvancedEditActivate
                    End If
                End If
            End If
        Case Else
            If m_bGridFocused Then
                If m_bItemActive Then
                    EditUpdateLabel
                    m_cSkinScrollBars.Refresh
                    bHandled = True
                End If
            Else
                lReturn = m_cHGridSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
            End If
        End Select
    
    '/* accelerators
    Case WM_CHAR
        '/* s - first row
        Select Case wParam
        Case 115, 83
            If (m_lRowCount > 0) Then
                RowEnsureVisible LBound(m_cGridItem)
            End If
        '/* e - last row
        Case 101, 69
            If (m_lRowCount > 0) Then
                RowEnsureVisible (Count - 1)
            End If
        End Select
        bHandled = True
        
    '/* row height
    Case WM_MEASUREITEM
        CopyMemory tMsItm, ByVal lParam, LenB(tMsItm)
        tMsItm.ItemHeight = m_lRowHeight
        CopyMemory ByVal lParam, tMsItm, LenB(tMsItm)
        bHandled = True
        
    '/* list mouse move
    Case WM_MOUSEMOVE
        If (lHwnd = m_lParentHwnd) Then
            TreeViewSizeWidth
        Else
            '/* cell hot tracking
            If m_bGridFocused Then
                If m_bCellHotTrack Then
                    CellTrack
                End If
                '/* cell info tips
                If m_bCellTips Then
                    If Not m_bItemActive Then
                        If Not m_bEditorLoaded Then
                            If Not m_bRowDragging Then
                                CellTipTrack
                            End If
                        End If
                    End If
                End If
            End If
            '/* header height user change
            If m_bHeaderSizable Then
                If (HeaderHitState = HdrHitVertSizer) Then
                    If Not m_bColumnSizingHorizontal Then
                        If Not m_bColumnDragging Then
                            If LeftKeyState Then
                                If m_bCellTips Then
                                    StopTipTimer
                                End If
                                If m_bItemActive Then
                                    EditUpdateLabel
                                End If
                                ColumnSizingVertical = True
                                CreateTransitionMask
                                ColumnSizeHeight
                            End If
                        End If
                    End If
                End If
            End If
        End If
    
    '/* mouse wheel
    Case WM_MOUSEWHEEL
        If m_bItemActive Then
            EditUpdateLabel
        End If
    
    '/* scrollbar click
    Case WM_NCLBUTTONDOWN
        If m_bItemActive Then
            EditUpdateLabel
        End If

    '/* control focus
    Case WM_SETFOCUS
        '/* tab without ipao
        If (lHwnd = m_lParentHwnd) Then
            If Not (m_cTreeView Is Nothing) Then
                SetFocus m_lTVHwnd
            Else
                SetFocus m_lHGHwnd
            End If
            bHandled = True
        End If

    '/* cursor change
    Case WM_SETCURSOR
        If m_bCustomCursors Then
            If (lHwnd = m_lParentHwnd) Then
                If TreeViewDividerHitTest Then
                    If (m_eTvControlAlignment > 1) Then
                        m_cSkinHeader.SetNSSizingCursor
                    Else
                        m_cSkinHeader.SetWESizingCursor
                    End If
                ElseIf m_bTreeViewSizing Then
                    If (m_eTvControlAlignment > 1) Then
                        m_cSkinHeader.SetNSSizingCursor
                    Else
                        m_cSkinHeader.SetWESizingCursor
                    End If
                Else
                    m_cSkinHeader.SetNormalCursor
                End If
            ElseIf (lHwnd = m_lHGHwnd) Then
                If m_bColumnDragging Then
                    m_cSkinHeader.SetDragCursor
                Else
                    m_cSkinHeader.SetNormalCursor
                End If
            Else
                If Not (m_cTreeView Is Nothing) Then
                    m_cSkinHeader.SetNormalCursor
                End If
            End If
            bHandled = True
        End If
        
    '/* size change
    Case WM_SIZE
        If m_bItemActive Then
            EditUpdateLabel
        End If
        If (lHwnd = m_lParentHwnd) Then
            Resize
        End If

    '/* column divider mark/tooltip callback
    Case WM_TIMER
        If m_bTipTracking Then
            ToolTipTrack
        Else
            HotDivider
        End If
    
    '/* react to settings changes
    Case WM_SETTINGCHANGE, WM_DISPLAYCHANGE
        UserControl_Resize
    
    '/* list notifications
    Case WM_NOTIFY
        CopyMemory tNmhdr, ByVal lParam, LenB(tNmhdr)
        '/* get msg code and owner
        With tNmhdr
            lCode = .code
            lOrigin = .hwndFrom
        End With
        '/*** header notifications ***/
        Select Case lOrigin
        Case m_lHdrHwnd
            Select Case lCode
            Case HDN_ITEMCHANGINGW, HDN_ITEMCHANGINGA
                If m_bItemActive Then
                    EditUpdateLabel
                End If
                CopyMemory tHDN, ByVal lParam, LenB(tHDN)
                If LeftKeyState Then
                    If m_bHeaderFixed Then
                        bHandled = True
                        lReturn = 1
                        Exit Sub
                    ElseIf m_bColumnLock(tHDN.iItem) Then
                        lReturn = 1
                        bHandled = True
                        Exit Sub
                    ElseIf m_bColumnDragLine Then
                        m_lColumnDivider = tHDN.iItem
                    End If
                    ColumnSizingHorizontal = True
                    m_cSkinScrollBars.Refresh
                Else
                    m_lColumnDivider = -1
                End If

            '/* header start drag
            Case HDN_BEGINDRAG
                If m_bItemActive Then
                    EditUpdateLabel
                End If
                CopyMemory tHDN, ByVal lParam, LenB(tHDN)
                If m_bDragDrop Then
                    If m_bColumnSizingHorizontal Then
                        lReturn = HDN_ENDDRAG
                    ElseIf m_bColumnSizingVertical Then
                        lReturn = HDN_ENDDRAG
                    ElseIf m_bFilterLoaded Then
                        lReturn = HDN_ENDDRAG
                    Else
                        If (HeaderHitTest = HdrHitColumn) Then
                            RaiseEvent eVHColumnDragging(tHDN.iItem)
                            ColumnDragging = True
                            m_lDraggedColumn = tHDN.iItem
                            ColumnIconStore tHDN.iItem, False
                            DragStartTimer
                        Else
                            lReturn = HDN_ENDDRAG
                        End If
                    End If
                Else
                    lReturn = HDN_ENDDRAG
                End If
                bHandled = True

            '/* post drag reset
            Case HDN_ENDDRAG
                HeaderDragStop
            
            '/* header horz size complete
            Case HDN_ENDTRACKA, HDN_ENDTRACKW
                If m_bColumnLock(tHDN.iItem) Then
                    bHandled = True
                End If
                If m_bColumnDragLine Then
                    m_lColumnDivider = -1
                    GridRefresh False
                End If
                RaiseEvent eVHColumnHorizontalSize(tHDN.iItem)
                ColumnSizingHorizontal = False
            End Select
        Case m_lHGHwnd
            '/*** list notifications ***/
            Select Case lCode
            '/* drag and drop trigger
            Case LVN_BEGINDRAG, LVN_BEGINRDRAG
                If m_bItemActive Then
                    EditUpdateLabel
                End If
                If m_bColumnSizingHorizontal Then
                    lReturn = LVN_ENDDRAG
                Else
                    If m_bCellTips Then
                        StopTipTimer
                    End If
                    CopyMemory tNmList, ByVal lParam, Len(tNmList)
                    m_bRowDragging = True
                    RaiseEvent eVHItemDragging(tNmList.iItem)
                    UserControl.OLEDrag
                End If
                bHandled = True
            
            '/* column click
            Case LVN_COLUMNCLICK
                If m_bItemActive Then
                    EditUpdateLabel
                End If
                CopyMemory tNmList, ByVal lParam, Len(tNmList)
                lColumn = tNmList.iSubItem
                Select Case HeaderHitState
                Case HdrHitFilter
                    If m_bColumnFilters Then
                        'If Not (m_lRowCount = 0) Then
                            FilterLoaded = True
                            FilterCreate lColumn
                        'End If
                        bHandled = True
                    End If
                Case HdrHitColumn
                    If m_bUseSorted Then
                        If Not (m_lRowCount = 0) Then
                            If Not m_bColumnSizingHorizontal Then
                                If Not m_bColumnSizingVertical Then
                                    RaiseEvent eVHColumnClick(lColumn)
                                    If m_bSkinHeader Then
                                        '/* swap od drawn sort icon
                                        With m_cSkinHeader
                                            If .ColumnSortDescending Then
                                                If .ColumnSorted = lColumn Then
                                                    bDesc = True
                                                    .ColumnSortDescending = False
                                                End If
                                            Else
                                                .ColumnSortDescending = True
                                            End If
                                            .ColumnSorted = lColumn
                                        End With
                                    Else
                                        '/* standard column icons
                                        ColumnIconReset
                                        If (ColumnIcon(lColumn) = -1) Then
                                            ColumnIcon(lColumn) = 1
                                        ElseIf (ColumnIcon(lColumn) = 1) Then
                                            ColumnIcon(lColumn) = 0
                                            bDesc = True
                                        Else
                                            ColumnIcon(lColumn) = 1
                                        End If
                                    End If
                                    '/* column and sort direction
                                    RaiseEvent eVHColumnClick(lColumn)
                                    If Not m_bVirtualMode Then
                                        If bDesc Then
                                            SortControl True, lColumn
                                        Else
                                            SortControl False, lColumn
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End Select
        
            '/* cell doubleclick
            Case NM_DBLCLK
                If Not CheckBoxHitTest Then
                    RaiseEvent eVHItemDoubleClick(RowInFocus, CellInFocus)
                    If m_bCellEdit Then
                        If Not (m_cGridItem(RowInFocus).RowNoEdit) Then
                            RaiseEvent eVHEditRequest(RowInFocus, CellInFocus)
                            EditBoxActivate
                        End If
                    End If
                End If
            
            '/* cell right-click
            Case NM_RCLICK
                Dim tPnt As POINTAPI
                GetCursorPos tPnt
                CopyMemory tNmList, ByVal lParam, Len(tNmList)
                RaiseEvent eVHItemRightClick(tNmList.iItem, tPnt.x, tPnt.y)
            
            '/* gained focus
            Case NM_SETFOCUS
                If m_bItemActive Then
                    EditUpdateLabel
                Else
                    If Not (m_lRowCount = 0) Then
                        If Not m_bGridFocused Then
                            If m_bFullRowSelect Then
                                CellRefresh RowInFocus, -1
                            Else
                                CellRefresh RowInFocus, CellInFocus
                            End If
                        End If
                    End If
                End If
                m_bGridFocused = True
                bHandled = True
            
            '/* lost focus
            Case NM_KILLFOCUS
                If Not m_bFilterLoaded Then
                    m_bGridFocused = False
                End If
                If Not (m_lRowCount = 0) Then
                    If m_bFullRowSelect Then
                        CellRefresh RowInFocus, -1
                    Else
                        CellRefresh RowInFocus, CellInFocus
                    End If
                End If
                StopTipTimer
            
            '/* works, but not in commctrl.h?
            Case LVN_ENDDRAG
                m_bColumnsShuffled = ColumnsAreShuffled
            End Select
            
        Case m_lTVHwnd
            Select Case lCode
            Case TVN_BEGINDRAGA, TVN_BEGINDRAGW
                If (m_lRowCount > 0) Then
                    CopyMemory uNMTV, ByVal lParam, Len(uNMTV)
                    m_lhNodeDrag = uNMTV.itemNew.hItem
                    If (m_lhNodeDrag > 0) Then
                        m_bNodeDragging = True
                        SendMessageLongA m_lTVHwnd, TVM_SELECTITEM, TVGN_CARET, m_lhNodeDrag
                        UserControl.OLEDrag
                    End If
                End If
            End Select
        End Select
    End Select

End Sub

Private Sub ColumnIconStore(ByVal lColumn As Long, _
                            ByVal bShow As Boolean)

'/* hide icon while dragging (black mask w/internal imagelist)
    
    If bShow Then
        If Not (m_lStoredIcon = -1) Then
            ColumnIcon(lColumn) = m_lStoredIcon
        End If
    Else
        m_lStoredIcon = ColumnIcon(lColumn)
        If Not (m_lStoredIcon = -1) Then
            ColumnIcon(lColumn) = -1
        End If
    End If
    
End Sub
                            
Private Sub CellKeyNavigate(ByVal eNavigate As EKNKeyNavigate)
'/* keyboard cell navigation

Dim lCount  As Long
Dim lRcnt   As Long
Dim lCell   As Long

    If (m_lRowCount = 0) Then Exit Sub
    
    lCount = (m_lColumnCount - 1)
    lRcnt = (m_lRowCount - 1)
    If m_bFullRowSelect Then
        lCell = -1
    Else
        lCell = CellInFocus
    End If
    Select Case eNavigate
    Case EKNRight
        With m_cGridItem(RowInFocus)
            If (CellInFocus = .SpanFirstCell) Then
                CellInFocus = .SpanLastCell
            End If
        End With
        If (CellInFocus < lCount) Then
            If CellIsSpanned(RowInFocus, CellInFocus) Then
                CellRefresh RowInFocus, lCell
            End If
            CellRefresh RowInFocus, CellInFocus
            CellInFocus = (CellInFocus + 1)
            CellScrollList RowInFocus, CellInFocus
            CellRefresh RowInFocus, CellInFocus
        Else
            If RowFocused(lRcnt) Then
                Exit Sub
            Else
                If CellIsSpanned(RowInFocus, CellInFocus) Then
                    CellRefresh RowInFocus, lCell
                End If
                CellRefresh RowInFocus, CellInFocus
                CellInFocus = 0
                RowInFocus = (RowInFocus + 1)
                CellScrollList RowInFocus, CellInFocus
                CellRefresh RowInFocus, CellInFocus
            End If
        End If
        
    Case EKNLeft
        If (CellInFocus = 0) Then
            If RowFocused(0) Then
                Exit Sub
            Else
                CellRefresh RowInFocus, CellInFocus
                CellInFocus = lCount
                RowInFocus = (RowInFocus - 1)
                CellScrollList RowInFocus, CellInFocus
                With m_cGridItem(RowInFocus)
                    If (CellInFocus = .SpanLastCell) Then
                        CellInFocus = .SpanFirstCell
                    End If
                End With
                CellRefresh RowInFocus, CellInFocus
            End If
        Else
            CellRefresh RowInFocus, CellInFocus
            CellInFocus = (CellInFocus - 1)
            With m_cGridItem(RowInFocus)
                If Not (.SpanFirstCell = -1) Then
                    If (CellInFocus = .SpanLastCell) Then
                        CellInFocus = .SpanFirstCell
                    End If
                End If
            End With
            CellScrollList RowInFocus, CellInFocus
            CellRefresh RowInFocus, CellInFocus
        End If
    
    Case EKNUp
        CellRefresh RowInFocus, lCell
        If (RowInFocus = 0) Then
            RowInFocus = 0
        Else
            RowInFocus = (RowInFocus - 1)
        End If
        If CellIsSpanned(RowInFocus, CellInFocus) Then
            CellInFocus = m_cGridItem(RowInFocus).SpanFirstCell
        End If
        CellScrollList RowInFocus, CellInFocus
        If m_bFullRowSelect Then
            CellRefresh RowInFocus, -1
        Else
            CellRefresh RowInFocus, CellInFocus
        End If
    
    Case EKNDown
        CellRefresh RowInFocus, lCell
        If (RowInFocus = lRcnt) Then
            RowInFocus = lRcnt
        Else
            RowInFocus = (RowInFocus + 1)
        End If
        CellScrollList RowInFocus, CellInFocus
        If m_bFullRowSelect Then
            CellRefresh RowInFocus, -1
        Else
            CellRefresh RowInFocus, CellInFocus
        End If
    End Select

End Sub

Private Sub CellScrollList(ByVal lRow As Long, _
                           ByVal lCell As Long)
'/* sync scrolled items into view

Dim lDepth  As Long
Dim lCt     As Long
Dim tRect   As RECT

    CellCalcRect lRow, lCell, tRect
    With tRect
        If LVHasVertical Then
            If m_bUseSpannedRows Then
                lDepth = m_cGridItem(lRow).SpanRowDepth
                If (lDepth < 1) Then
                    lDepth = 1
                End If
                For lCt = 1 To lDepth
                    If (.top < (m_lHeaderOffset)) Then
                        LVScrollVertical False
                    ElseIf (.Bottom > (ScaleHeight / Screen.TwipsPerPixelY)) Then
                        LVScrollVertical True
                    End If
                Next lCt
            Else
                If (.top < (m_lHeaderOffset)) Then
                    LVScrollVertical False
                ElseIf (.Bottom > (ScaleHeight / Screen.TwipsPerPixelY)) Then
                    LVScrollVertical True
                End If
            End If
        End If
    End With

End Sub


'**********************************************************************
'*                              DRAG AND DROP
'**********************************************************************

Private Sub UserControl_OLECompleteDrag(Effect As Long)
'/* drag completed

    m_cDragImage.CompleteDrag
    If (m_eDragEffectStyle = edsClientArrow) Then
        DragLineReset
    End If
    m_bRowDragging = False
    m_bNodeDragging = False

End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, _
                                    Effect As Long, _
                                    Button As Integer, _
                                    Shift As Integer, _
                                    x As Single, _
                                    y As Single)

'/* complete drag operation

Dim lIndex  As Long
Dim tLVHT   As LVHITTESTINFO
Dim tPoint  As POINTAPI

    '/* get target item index
    If ((Effect And vbDropEffectMove) = vbDropEffectMove) Then
        GetCursorPos tPoint
        ScreenToClient m_lHGHwnd, tPoint
        lIndex = -1
        LSet tLVHT.pt = tPoint
        SendMessageA m_lHGHwnd, LVM_HITTEST, 0&, tLVHT
        If (tLVHT.iItem <= 0) Then
            If (tLVHT.flags And LVHT_NOWHERE) = LVHT_NOWHERE Then
                lIndex = FindNearestItem(tPoint)
            Else
                lIndex = tLVHT.iItem
            End If
        Else
            lIndex = tLVHT.iItem
        End If
        m_cDragImage.CompleteDrag
    End If

    If m_bNodeDragging Then
        If (Len(Data.GetData(vbCFText)) > 0) Then
            If Not (m_lLastRow = -1) Then
                If Not (m_lLastCell = -1) Then
                    CellText(m_lLastRow, m_lLastCell) = Data.GetData(vbCFText)
                End If
            End If
        End If
    Else
        If Not m_bVirtualMode Then
            If m_bUseSpannedRows Then
                lIndex = RowSpanMapVirtual(lIndex)
            End If
            '/* swap items
            If (lIndex > -1) Then
                If (RowInFocus > -1) Then
                    MoveArrayItem RowInFocus, lIndex
                    If m_bHasSubCells Then
                        SubCellSortRows
                    End If
                End If
            End If
        End If
    End If
    
    '/* select cell
    RowSelected(RowInFocus) = False
    GridRefresh False
    RowFocused(lIndex) = True
    RaiseEvent eVHItemDragComplete(RowInFocus, lIndex)
    '/* refresh
    If Not (m_cSkinScrollBars Is Nothing) Then
        m_cSkinScrollBars.Refresh
    End If

End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, _
                                    Effect As Long, _
                                    Button As Integer, _
                                    Shift As Integer, _
                                    x As Single, _
                                    y As Single, _
                                    State As Integer)

'/* move scrollbars during drag

Dim lIndex      As Long
Dim tPoint      As POINTAPI
Dim tRect       As RECT

    '/* scroll list when required
    GetCursorPos tPoint
    GetWindowRect m_lHGHwnd, tRect
    '/* vertical scroll
    With m_cDragImage
        If LVHasVertical Then
            If (Abs(tPoint.y - tRect.top) < (m_lHeaderHeight + 6)) Then
                .HideDragImage True
                LVScrollVertical False
                .HideDragImage False
            ElseIf (Abs(tPoint.y - tRect.Bottom) < (m_lRowHeight + 6)) Then
                .HideDragImage True
                LVScrollVertical True
                .HideDragImage False
            End If
        End If
        '/* horizontal scroll
        If LVHasHorizontal Then
            If (Abs(tPoint.x - tRect.left) < 24) Then
                .HideDragImage True
                LVScrollHorizontal False
                .HideDragImage False
            ElseIf (Abs(tPoint.x - tRect.Right) < 48) Then
                .HideDragImage True
                LVScrollHorizontal True
                .HideDragImage False
            End If
        End If
    End With
    '/* highlite drag over items
    If Not m_bNodeDragging Then
        lIndex = OverItem
        If Not (lIndex = -1) Then
            If Not (lIndex = m_lLastDropIdx) Then
                DrawOleDragLine lIndex
                m_lLastDropIdx = lIndex
            End If
        End If
    End If
    
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, _
                                        DefaultCursors As Boolean)

'/* refresh drag image
    
    m_cDragImage.DragDrop
    If m_bNodeDragging Then
        CellHilite
    End If

End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, _
                                     AllowedEffects As Long)

'/* initialize drag operation

Dim bData() As Byte
Dim lIcHnd  As Long
Dim sData   As String

    If Not m_bHeaderSizing Then
        If Not (OLEDropMode = vbOLEDropNone) Then
            AllowedEffects = vbDropEffectMove Or vbDropEffectCopy
            If m_bNodeDragging Then
                sData = TreeViewNodeText(m_lhNodeDrag)
            End If
            bData = sData
            With Data
                .Clear
                .SetData bData, vbCFText
            End With
            CreateDragImage
            lIcHnd = m_lDragImgIml
            With m_cDragImage
                .Parent = UserControl.hwnd
                .hImageList = lIcHnd
                If m_bNodeDragging Then
                    .StartDrag 0, -8, -30
                Else
                    .StartDrag 0, -8, -8
                End If
            End With
        End If
    End If

End Sub

Private Sub CellHilite()
'/* hilite drag candidate

Dim lCell   As Long
Dim lHdc    As Long
Dim lRow    As Long
Dim tRect   As RECT
Dim tROft   As RECT

    lRow = RowHitTest
    lCell = CellHitTest
    If (lRow = m_lLastRow) Then
        If (lCell = m_lLastCell) Then
            Exit Sub
        End If
    End If
    m_cDragImage.HideDragImage True
    CellRefresh m_lLastRow, m_lLastCell
    RePaint
    If Not (lRow = -1) Then
        If Not (lCell = -1) Then
            lHdc = GetDC(m_lHGHwnd)
            CellCalcRect lRow, lCell, tRect
            CopyRect tROft, tRect
            With tRect
                Select Case m_eCellHiliteStyle
                Case echStripe
                    InflateRect tROft, -4, -4
                    With tROft
                        ExcludeClipRect lHdc, .left, .top, .Right, .Bottom
                    End With
                    m_cRender.FramePattern lHdc, .left, .top, .Right, .Bottom, 4, epsDot, m_lCellHiliteColor
                Case echThin
                    InflateRect tROft, -1, -1
                    With tROft
                        ExcludeClipRect lHdc, .left, .top, .Right, .Bottom
                    End With
                    m_cRender.FramePattern lHdc, .left, .top, .Right, .Bottom, 1, epsSolid, m_lCellHiliteColor
                Case echThick
                    InflateRect tROft, -2, -2
                    With tROft
                        ExcludeClipRect lHdc, .left, .top, .Right, .Bottom
                    End With
                    m_cRender.FramePattern lHdc, .left, .top, .Right, .Bottom, 2, epsSolid, m_lCellHiliteColor
                End Select
            End With
            ReleaseDC m_lHGHwnd, lHdc
        End If
    End If
    m_cDragImage.HideDragImage False
    m_lLastRow = lRow
    m_lLastCell = lCell
    
End Sub

Private Sub CreateDragImage()
'/* create row drag image

Dim lHdc    As Long
Dim lHwnd   As Long
Dim lhBmp   As Long
Dim tRClt   As RECT
Dim tRRow   As RECT
Dim cDragDc As clsStoreDc

    If m_bNodeDragging Then
        GetClientRect m_lTVHwnd, tRClt
        NodeGetRect m_lhNodeDrag, tRRow
        If (tRClt.Right < tRRow.Right) Then
            tRRow.Right = tRClt.Right
        End If
        InflateRect tRRow, 0, -1
        lHwnd = m_lTVHwnd
        lHdc = GetDC(lHwnd)
    Else
        GetClientRect m_lHGHwnd, tRClt
        CellCalcRect RowInFocus, -1, tRRow
        tRRow.Right = tRClt.Right
        InflateRect tRRow, 0, 1
        lHwnd = m_lHGHwnd
        lHdc = GetDC(lHwnd)
    End If
    Set cDragDc = New clsStoreDc

    With tRRow
        cDragDc.Width = (.Right - .left)
        cDragDc.Height = (.Bottom - .top)
        m_cRender.Blit cDragDc.hdc, 0, 0, (.Right - .left), (.Bottom - .top), lHdc, .left, .top, SRCCOPY
        lhBmp = cDragDc.ExtractBitmap
        InitImlDrag (.Right - .left), (.Bottom - .top)
    End With
    
    ImlDragAddBmp lhBmp
    Set cDragDc = Nothing
    ReleaseDC lHwnd, lHdc

End Sub

Private Sub DrawOleDragLine(ByVal lRow As Long)
'/* draw ole row drag marks

Dim lPrHnd      As Long
Dim lParDc      As Long
Dim lGrdDc      As Long
Dim lXPos       As Long
Dim lYPos       As Long
Dim lhPen       As Long
Dim lhPenOld    As Long
Dim tPnt        As POINTAPI
Dim tPcd        As POINTAPI
Dim tRect       As RECT
Dim tRClt       As RECT

    CellCalcRect lRow, -1, tRClt
    DragLineReset
    If (tRClt.Bottom > (UserControl.ScaleHeight / Screen.TwipsPerPixelY)) Then
        Exit Sub
    End If
    If (m_eDragEffectStyle = edsClientArrow) Then
        m_lRowDragEffect = lRow
        '/* get handle/sizes
        lPrHnd = GetParent(m_lParentHwnd)
        If lPrHnd = 0 Then Exit Sub
        lParDc = GetDC(lPrHnd)
        GetWindowRect m_lParentHwnd, tRect
        CopyMemory tPcd, tRect, Len(tPcd)
        ScreenToClient lPrHnd, tPcd
        '/* left arrow
        With tPcd
            lYPos = (.y + tRClt.Bottom)
            lXPos = (.x - 13)
        End With
        '/* left arrow line
        lhPen = CreatePen(0&, 2&, 255&)
        lhPenOld = SelectObject(lParDc, lhPen)
        MoveToEx lParDc, lXPos, lYPos - 1, tPnt
        LineTo lParDc, lXPos + 6, lYPos
        SelectObject lParDc, lhPenOld
        DeleteObject lhPen
        lhPenOld = 0
        lhPen = 0
        '/* left arrowhead
        lhPen = CreatePen(0&, 1&, 255&)
        lhPenOld = SelectObject(lParDc, lhPen)
        MoveToEx lParDc, (lXPos + 7), (lYPos - 5), tPnt
        LineTo lParDc, (lXPos + 7), (lYPos + 5)
        MoveToEx lParDc, (lXPos + 8), (lYPos - 4), tPnt
        LineTo lParDc, (lXPos + 8), (lYPos + 4)
        MoveToEx lParDc, (lXPos + 9), (lYPos - 3), tPnt
        LineTo lParDc, (lXPos + 9), (lYPos + 3)
        MoveToEx lParDc, (lXPos + 10), (lYPos - 2), tPnt
        LineTo lParDc, (lXPos + 10), (lYPos + 2)
        MoveToEx lParDc, (lXPos + 11), (lYPos - 1), tPnt
        LineTo lParDc, (lXPos + 11), (lYPos + 1)
        SelectObject lParDc, lhPenOld
        DeleteObject lhPen
        lhPenOld = 0
        lhPen = 0
        '/* store coordinates
        With m_tDividerRect(0)
            .top = lYPos - 5
            .Bottom = lYPos + 6
            .left = lXPos
            .Right = lXPos + 12
        End With
        '/* right arrow
        With tRect
            lYPos = tPcd.y + tRClt.Bottom
            lXPos = (.Right - .left) + (tPcd.x + 18)
        End With
        '/* arrow line
        lhPen = CreatePen(0&, 2&, 255&)
        lhPenOld = SelectObject(lParDc, lhPen)
        MoveToEx lParDc, lXPos - 13, lYPos - 1, tPnt
        LineTo lParDc, lXPos - 6, lYPos
        SelectObject lParDc, lhPenOld
        DeleteObject lhPen
        lhPenOld = 0
        lhPen = 0
        '/* right arrowhead
        lhPen = CreatePen(0&, 1&, 255&)
        lhPenOld = SelectObject(lParDc, lhPen)
        MoveToEx lParDc, (lXPos - 14), (lYPos - 5), tPnt
        LineTo lParDc, (lXPos - 14), (lYPos + 5)
        MoveToEx lParDc, (lXPos - 15), (lYPos - 4), tPnt
        LineTo lParDc, (lXPos - 15), (lYPos + 4)
        MoveToEx lParDc, (lXPos - 16), (lYPos - 3), tPnt
        LineTo lParDc, (lXPos - 16), (lYPos + 3)
        MoveToEx lParDc, (lXPos - 17), (lYPos - 2), tPnt
        LineTo lParDc, (lXPos - 17), (lYPos + 2)
        MoveToEx lParDc, (lXPos - 18), (lYPos - 1), tPnt
        LineTo lParDc, (lXPos - 18), (lYPos + 1)
        SelectObject lParDc, lhPenOld
        DeleteObject lhPen
        lhPenOld = 0
        lhPen = 0
        With m_tDividerRect(1)
            .top = lYPos - 6
            .Bottom = lYPos + 6
            .left = lXPos - 19
            .Right = lXPos - 5
        End With
        '/* cleanup
        SelectObject lParDc, lhPenOld
        DeleteObject lhPen
        lhPenOld = 0
        lhPen = 0
        ReleaseDC lPrHnd, lParDc
    Else
        '/* draw corner bracket
        lGrdDc = GetDC(m_lHGHwnd)
        If (m_eDragEffectStyle = edsThinLine) Then
            lhPen = CreatePen(0&, 1&, 255&)
        Else
            lhPen = CreatePen(0&, 2&, 255&)
        End If
        lhPenOld = SelectObject(lGrdDc, lhPen)
        With tRClt
            MoveToEx lGrdDc, .left, (.Bottom - 12), tPnt
            LineTo lGrdDc, .left, (.Bottom - 2)
            MoveToEx lGrdDc, .left, (.Bottom - 2), tPnt
            LineTo lGrdDc, (.left + 12), (.Bottom - 2)
        End With
        '/* store coordinates
        CopyRect m_tDividerRect(0), tRClt
        SelectObject lGrdDc, lhPenOld
        DeleteObject lhPen
        lhPenOld = 0
        lhPen = 0
        ReleaseDC m_lHGHwnd, lGrdDc
    End If

End Sub

Private Sub DragLineReset()
'/* reset drag marks

Dim lPrHnd  As Long

    If (m_eDragEffectStyle = edsClientArrow) Then
        lPrHnd = GetParent(m_lParentHwnd)
        EraseRect lPrHnd, m_tDividerRect(0), 0&
        If lPrHnd = 0 Then Exit Sub
        EraseRect lPrHnd, m_tDividerRect(1), 0&
    Else
        EraseRect m_lHGHwnd, m_tDividerRect(0), 0&
    End If

End Sub

Private Function FindNearestItem(ByRef tPoint As POINTAPI) As Long
'/* return closest item index

Dim lX          As Long
Dim lY          As Long
Dim lCt         As Long
Dim lCount      As Long
Dim lDistSq     As Long
Dim lMinDistSq  As Long
Dim lMinItem    As Long
Dim tRect       As RECT

    lMinItem = -1
    lMinDistSq = &H7FFFFFFF
    lCount = RowCount
    For lCt = 1 To lCount
        tRect.left = LVIR_BOUNDS
        SendMessageA m_lHGHwnd, LVM_GETITEMRECT, lCt - 1, tRect
        With tRect
            lX = tPoint.x - (.left + (.Right - .left) \ 2)
            lY = tPoint.y - (.top + (.Bottom - .top) \ 2)
        End With
        lDistSq = (lX * lX) + (lY * lY)
        If (lDistSq < lMinDistSq) Then
            lMinDistSq = lDistSq
            lMinItem = lCt
            Exit For
        End If
    Next lCt
    FindNearestItem = lMinItem

End Function

Private Function LVScrollVertical(ByVal bDown As Boolean)
'/* scroll vertical

    If bDown Then
        SendMessageLongA m_lHGHwnd, WM_VSCROLL, SB_LINEDOWN, 0
    Else
        SendMessageLongA m_lHGHwnd, WM_VSCROLL, SB_LINEUP, 0
    End If

End Function

Private Function LVScrollHorizontal(ByVal bRight As Boolean)
'/* scroll horizontal

    If bRight Then
        SendMessageLongA m_lHGHwnd, WM_HSCROLL, SB_LINERIGHT, 0
    Else
        SendMessageLongA m_lHGHwnd, WM_HSCROLL, SB_LINELEFT, 0
    End If

End Function

Private Function LVHasHorizontal() As Boolean
'/* vertical scrollbar test

Dim lStyle  As Long

    If m_bIsNt Then
        lStyle = GetWindowLongW(m_lHGHwnd, GWL_STYLE)
    Else
        lStyle = GetWindowLongA(m_lHGHwnd, GWL_STYLE)
    End If
    LVHasHorizontal = (lStyle And WS_HSCROLL) <> 0

End Function

Private Function LVHasVertical() As Boolean
'/* horizontal scrollbar test

Dim lStyle  As Long

    If m_bIsNt Then
        lStyle = GetWindowLongW(m_lHGHwnd, GWL_STYLE)
    Else
        lStyle = GetWindowLongA(m_lHGHwnd, GWL_STYLE)
    End If
    LVHasVertical = (lStyle And WS_VSCROLL) <> 0

End Function

Private Sub NodeGetRect(ByVal lNode As Long, _
                        tRect As RECT)

    tRect.left = lNode
    SendMessageA m_lTVHwnd, TVM_GETITEMRECT, True, tRect

End Sub

Private Function OverItem() As Long
'/* pointer over item

Dim lIndex  As Long
Dim tLVHT   As LVHITTESTINFO
Dim tPoint  As POINTAPI

    '/* get target item index
    GetCursorPos tPoint
    ScreenToClient m_lHGHwnd, tPoint
    lIndex = -1
    LSet tLVHT.pt = tPoint
    SendMessageA m_lHGHwnd, LVM_HITTEST, 0&, tLVHT
    If (tLVHT.iItem <= 0) Then
        If (tLVHT.flags And LVHT_NOWHERE) = LVHT_NOWHERE Then
            lIndex = -1
        Else
            lIndex = tLVHT.iItem
        End If
    Else
        lIndex = tLVHT.iItem
    End If
    If m_bUseSpannedRows Then
        lIndex = RowSpanMapVirtual(lIndex)
    End If
    OverItem = lIndex

End Function

Public Function RowIsVisible(ByVal lRow As Long) As Boolean
'/* return row in client area

Dim tRClt   As RECT
Dim tRRow   As RECT
Dim tRcpy   As RECT
Dim tPnt    As POINTAPI

    GetClientRect m_lHGHwnd, tRClt
    CellCalcRect lRow, -1, tRRow
    CopyRect tRcpy, tRRow
    tRRow.Bottom = tRRow.top + 6
    
    CopyMemory tPnt, tRRow, Len(tPnt)
    If Not (PtInRect(tRClt, tPnt.x, tPnt.y) = 0) Then
        tRcpy.top = tRcpy.Bottom - 6
        CopyMemory tPnt, tRcpy, Len(tPnt)
        If Not (PtInRect(tRClt, tPnt.x, tPnt.y) = 0) Then
            RowIsVisible = True
        End If
    End If

End Function


'**********************************************************************
'*                              CONTROL
'**********************************************************************

Private Sub UserControl_Hide()
'/* hide grid

    If m_bSkinScrollBars Then
        If Not (m_cSkinScrollBars Is Nothing) Then
            m_cSkinScrollBars.Visible = False
        End If
    End If
    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.Visible = False
    End If
    m_bVisible = False

End Sub

Private Sub UserControl_Paint()
'/* draw caption

    If (m_lHGHwnd = 0) Then
        If Not (lblName.Caption = UserControl.Extender.Name) Then
            lblName.Caption = UserControl.Extender.Name
        End If
    End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'/* read properties

Dim bIcons()    As Byte
Dim bKeys()     As Byte
Dim sFont       As New StdFont

    Initialize
    With PropBag
        Set Font = .ReadProperty("Font", sFont)
        Set HeaderFont = .ReadProperty("HeaderFont", sFont)
        Set UserControl.Font = Font
        If Not (m_cHeaderIcons Is Nothing) Then
            ReDim bIcons(0)
            ReDim bKeys(0)
            m_cHeaderIcons.IconSizeX = .ReadProperty("HeaderIconSizeX", 16)
            m_cHeaderIcons.IconSizeY = .ReadProperty("HeaderIconSizeY", 16)
            m_cHeaderIcons.ColourDepth = .ReadProperty("HeaderIconColourDepth", &H18)
            m_cHeaderIcons.ImageCount = .ReadProperty("HeaderIconCount", 0)
            On Error Resume Next
            bIcons = .ReadProperty("HeaderIcons", "")
            bKeys = .ReadProperty("HeaderKeys", "")
            If (UBound(bIcons) > 0) Then
                m_cHeaderIcons.RestoreIcons bKeys, bIcons
            End If
            On Error GoTo 0
        End If
        If Not (m_cCellIcons Is Nothing) Then
            ReDim bIcons(0)
            ReDim bKeys(0)
            m_cCellIcons.IconSizeX = .ReadProperty("CellIconSizeX", 16)
            m_cCellIcons.IconSizeY = .ReadProperty("CellIconSizeY", 16)
            m_cCellIcons.ColourDepth = .ReadProperty("CellIconColourDepth", &H18)
            m_cCellIcons.ImageCount = .ReadProperty("CellIconCount", 0)
            On Error Resume Next
            bIcons = .ReadProperty("CellIcons", "")
            bKeys = .ReadProperty("CellKeys", "")
            If (UBound(bIcons) > 0) Then
                m_cCellIcons.RestoreIcons bKeys, bIcons
            End If
            On Error GoTo 0
        End If
        If Not (m_cTreeIcons Is Nothing) Then
            ReDim bIcons(0)
            ReDim bKeys(0)
            m_cTreeIcons.IconSizeX = .ReadProperty("TreeIconSizeX", 16)
            m_cTreeIcons.IconSizeY = .ReadProperty("TreeIconSizeY", 16)
            m_cTreeIcons.ColourDepth = .ReadProperty("TreeIconColourDepth", &H18)
            m_cTreeIcons.ImageCount = .ReadProperty("TreeIconCount", 0)
            On Error Resume Next
            bIcons = .ReadProperty("TreeIcons", "")
            bKeys = .ReadProperty("TreeKeys", "")
            If (UBound(bIcons) > 0) Then
                m_cTreeIcons.RestoreIcons bKeys, bIcons
            End If
            On Error GoTo 0
        End If
        AlphaBarTransparency = .ReadProperty("AlphaBarTransparency", PRP_APT)
        AlphaBarTheme = .ReadProperty("AlphaBarTheme", False)
        AlphaBarActive = .ReadProperty("AlphaBarActive", False)
        BackColor = .ReadProperty("BackColor", &HFFFFFF)
        BorderStyle = .ReadProperty("BorderStyle", PRP_BRDSTL)
        CellEdit = .ReadProperty("CellEdit", False)
        CellFocusedColor = .ReadProperty("CellFocusedColor", GetSysColor(vbHighlight And &H1F&))
        CellSelectedColor = .ReadProperty("CellSelectedColor", GetSysColor(vbButtonFace And &H1F&))
        CheckBoxes = .ReadProperty("Checkboxes", False)
        ColumnDragLine = .ReadProperty("ColumnDragLine", False)
        ColumnFocus = .ReadProperty("ColumnFocus", False)
        ColumnFocusColor = .ReadProperty("ColumnFocusColor", &H303030)
        CustomCursors = .ReadProperty("CustomCursors", False)
        DoubleBuffer = .ReadProperty("DoubleBuffer", False)
        DragEffectStyle = .ReadProperty("DragEffectStyle", edsClientArrow)
        Enabled = .ReadProperty("Enabled", True)
        FocusAlphaBlend = .ReadProperty("FocusAlphaBlend", False)
        FocusTextOnly = .ReadProperty("FocusTextOnly", False)
        ForeColor = .ReadProperty("ForeColor", vbWindowText)
        ForeColorFocused = .ReadProperty("ForeColorFocused", vbWhite)
        FullRowSelect = .ReadProperty("FullRowSelect", False)
        GridLines = .ReadProperty("GridLines", EGLBoth)
        GridLineColor = .ReadProperty("GridLineColor", GetSysColor(vbButtonShadow And &H1F&))
        HeaderDragDrop = .ReadProperty("HeaderDragDrop", True)
        HeaderFixedWidth = .ReadProperty("HeaderFixedWidth", False)
        HeaderForeColor = .ReadProperty("HeaderForeColor", &H404040)
        HeaderForeColorFocused = .ReadProperty("HeaderForeColorFocused", &H808080)
        HeaderForeColorPressed = .ReadProperty("HeaderForeColorPressed", &H202020)
        HeaderHeight = .ReadProperty("HeaderHeight", 28)
        HeaderHeightSizable = .ReadProperty("HeaderHeightSizable", True)
        OLEDragMode = .ReadProperty("OLEDragMode", vbOLEDragManual)
        OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
        SortType = .ReadProperty("SortType", estCaseSensitive)
        VirtualMode = .ReadProperty("VirtualMode", False)
        XPColors = .ReadProperty("XPColors", False)
    End With

End Sub

Private Sub UserControl_Resize()
'/* resize grid
    Resize
End Sub

Private Sub UserControl_Show()
'/* show control

    If Not (m_cTreeView Is Nothing) Then
        m_cTreeView.Visible = True
    End If
    If m_bSkinScrollBars Then
        If Not (m_cSkinScrollBars Is Nothing) Then
            If Not (m_cSkinScrollBars.Visible) Then
                m_cSkinScrollBars.Visible = True
                UserControl_Resize
                m_cSkinScrollBars.Refresh
                HeaderHeight = m_lHeaderHeight
                m_cSkinHeader.Refresh -1
            End If
        End If
    End If
    m_bVisible = True

End Sub

Private Sub UserControl_Terminate()
'/* destroy grid

    GridDetatch
    DragStopTimer
    DestroyItems
    DestroyList
    DestroyImages
    DeAllocatePointer "a", True
    Set m_cHGridSubclass = Nothing
    If Not (m_lhMod = 0) Then
        FreeLibrary m_lhMod
        m_lhMod = 0
    End If
    DestroyClasses
    m_lParentHwnd = 0
    m_lHGHwnd = 0
    m_lEditHwnd = 0

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'/* write properties

Dim sFont       As New StdFont
Dim bIcons()    As Byte
Dim bKeys()     As Byte

    With PropBag
        .WriteProperty "Font", Font, sFont
        If Not (m_cHeaderIcons Is Nothing) Then
            .WriteProperty "HeaderIconSizeX", m_cHeaderIcons.IconSizeX, 16
            .WriteProperty "HeaderIconSizeY", m_cHeaderIcons.IconSizeY, 16
            .WriteProperty "HeaderIconColourDepth", m_cHeaderIcons.ColourDepth, &H18
            .WriteProperty "HeaderIconCount", m_cHeaderIcons.ImageCount, 0
            ReDim bIcons(0)
            ReDim bKeys(0)
            On Error Resume Next
            m_cHeaderIcons.SaveIcons bKeys, bIcons
            If (UBound(bIcons) > 0) Then
                .WriteProperty "HeaderIcons", bIcons, ""
            End If
            If (UBound(bKeys) > 0) Then
                .WriteProperty "HeaderKeys", bKeys, ""
            End If
            On Error GoTo 0
        End If
        If Not (m_cCellIcons Is Nothing) Then
            .WriteProperty "CellIconSizeX", m_cCellIcons.IconSizeX, 16
            .WriteProperty "CellIconSizeY", m_cCellIcons.IconSizeY, 16
            .WriteProperty "CellIconColourDepth", m_cCellIcons.ColourDepth, &H18
            .WriteProperty "CellIconCount", m_cCellIcons.ImageCount, 0
            ReDim bIcons(0)
            ReDim bKeys(0)
            On Error Resume Next
            m_cCellIcons.SaveIcons bKeys, bIcons
            If (UBound(bIcons) > 0) Then
                .WriteProperty "CellIcons", bIcons, ""
            End If
            If (UBound(bKeys) > 0) Then
                .WriteProperty "CellKeys", bKeys, ""
            End If
            On Error GoTo 0
        End If
        If Not (m_cTreeIcons Is Nothing) Then
            .WriteProperty "TreeIconSizeX", m_cTreeIcons.IconSizeX, 16
            .WriteProperty "TreeIconSizeY", m_cTreeIcons.IconSizeY, 16
            .WriteProperty "TreeIconColourDepth", m_cTreeIcons.ColourDepth, &H18
            .WriteProperty "TreeIconCount", m_cTreeIcons.ImageCount, 0
            ReDim bIcons(0)
            ReDim bKeys(0)
            On Error Resume Next
            m_cTreeIcons.SaveIcons bKeys, bIcons
            If (UBound(bIcons) > 0) Then
                .WriteProperty "TreeIcons", bIcons, ""
            End If
            If (UBound(bKeys) > 0) Then
                .WriteProperty "TreeKeys", bKeys, ""
            End If
            On Error GoTo 0
        End If
        .WriteProperty "AlphaBarTheme", AlphaBarTheme, False
        .WriteProperty "AlphaBarTransparency", AlphaBarTransparency, PRP_APT
        .WriteProperty "AlphaBarActive", AlphaBarActive, False
        .WriteProperty "BackColor", BackColor, &HFFFFFF
        .WriteProperty "BorderStyle", BorderStyle, PRP_BRDSTL
        .WriteProperty "CellEdit", CellEdit, False
        .WriteProperty "CellFocusedColor", CellFocusedColor, GetSysColor(vbHighlight And &H1F&)
        .WriteProperty "CellSelectedColor", CellSelectedColor, GetSysColor(vbButtonFace And &H1F&)
        .WriteProperty "Checkboxes", CheckBoxes, False
        .WriteProperty "ColumnDragLine", ColumnDragLine, False
        .WriteProperty "ColumnFocus", ColumnFocus, False
        .WriteProperty "ColumnFocusColor", ColumnFocusColor, &H303030
        .WriteProperty "CustomCursors", CustomCursors, False
        .WriteProperty "DoubleBuffer", DoubleBuffer, False
        .WriteProperty "DragEffectStyle", DragEffectStyle, edsClientArrow
        .WriteProperty "Enabled", Enabled, True
        .WriteProperty "ForeColor", ForeColor, vbWindowText
        .WriteProperty "ForeColorFocused", ForeColorFocused, vbWhite
        .WriteProperty "FocusAlphaBlend", FocusAlphaBlend, False
        .WriteProperty "FocusTextOnly", FocusTextOnly, False
        .WriteProperty "FullRowSelect", FullRowSelect, False
        .WriteProperty "GridLines", GridLines, EGLBoth
        .WriteProperty "GridLineColor", GridLineColor, GetSysColor(vbButtonShadow And &H1F&)
        .WriteProperty "HeaderDragDrop", HeaderDragDrop, True
        .WriteProperty "HeaderFont", HeaderFont, sFont
        .WriteProperty "HeaderFixedWidth", HeaderFixedWidth, False
        .WriteProperty "HeaderForeColor", HeaderForeColor, &H404040
        .WriteProperty "HeaderForeColorFocused", HeaderForeColorFocused, &H808080
        .WriteProperty "HeaderForeColorPressed", HeaderForeColorPressed, &H202020
        .WriteProperty "HeaderHeight", HeaderHeight, 28
        .WriteProperty "HeaderHeightSizable", HeaderHeightSizable, False
        .WriteProperty "OLEDragMode", OLEDragMode, vbOLEDragManual
        .WriteProperty "OLEDropMode", OLEDropMode, vbOLEDropNone
        .WriteProperty "SortType", SortType, estCaseSensitive
        .WriteProperty "VirtualMode", VirtualMode, False
        .WriteProperty "XPColors", XPColors, False
    End With

End Sub


'**********************************************************************
'*                              CLEANUP
'**********************************************************************

Public Sub ClearList()
'/* clear all items

    DestroyItems
    DeAllocatePointer "a", True
    SetRowCount 0
    m_lRowCount = 0
    m_bHasInitialized = False
    RaiseEvent eVHItemClear

End Sub

Private Sub DeAllocatePointer(ByVal sKey As String, _
                              Optional ByVal bPurge As Boolean)

'/* purge memory pointers

Dim lPtr As Long
Dim lC   As Long

    If Not (c_PtrMem Is Nothing) Then
        If Not bPurge Then
            '/* get the pointer
            On Error Resume Next
            lPtr = c_PtrMem.Item(sKey)
            On Error GoTo 0
            If lPtr = 0 Then GoTo Handler
            '/* release the memory
            CopyMemory ByVal lPtr, 0&, 4&
        Else
            '/* destroy the struct last
            With c_PtrMem
                For lC = .Count To 1 Step -1
                    If Not (CLng(.Item(lC)) = 0) Then
                        lPtr = CLng(.Item(lC))
                        CopyMemory ByVal lPtr, 0&, 4&
                    End If
                Next lC
            End With
            m_lStrctPtr = 0
        End If
    End If

Handler:

End Sub

Private Sub DestroyArrays()
'*/ array cleanup

On Error Resume Next

    Erase m_oCellFont
    Erase m_tDividerRect
    Erase m_bColumnLock
    Erase m_hFontHnd
    Erase m_sSortArray
    Erase m_lCellColor
    Erase m_lSortArray

On Error GoTo 0

End Sub

Private Function DestroyFont() As Boolean
'*/ font cleanup

    If Not (m_lFont = 0) Then
        If DeleteObject(m_lFont) Then
            DestroyFont = True
            m_lFont = 0
        End If
    End If

End Function

Private Sub DestroyFonts()
'/* destroy grid fonts

Dim lCt As Long

    If ArrayCheck(m_oCellFont) Then
        For lCt = LBound(m_oCellFont) To UBound(m_oCellFont)
            If Not (m_oCellFont(lCt) Is Nothing) Then
                Set m_oCellFont(lCt) = Nothing
            End If
            If Not (m_hFontHnd(lCt) = 0) Then
                DeleteObject m_hFontHnd(lCt)
                m_hFontHnd(lCt) = 0
            End If
        Next lCt
    End If
    If Not m_oFont Is Nothing Then Set m_oFont = Nothing
    If Not m_oCellTipFont Is Nothing Then Set m_oCellTipFont = Nothing
    If Not m_oHeaderFont Is Nothing Then Set m_oHeaderFont = Nothing

End Sub

Private Sub DestroyImages()
'/* destroy images

    If Not m_IChecked Is Nothing Then Set m_IChecked = Nothing
    If Not m_IUnChecked Is Nothing Then Set m_IUnChecked = Nothing
    If Not m_IChkDisabled Is Nothing Then Set m_IChkDisabled = Nothing
    If Not m_cChkCheckDc Is Nothing Then Set m_cChkCheckDc = Nothing
    If Not m_cChkUnCheckDc Is Nothing Then Set m_cChkUnCheckDc = Nothing
    If Not m_cChkDisableDc Is Nothing Then Set m_cChkDisableDc = Nothing
    If Not m_cRender Is Nothing Then Set m_cRender = Nothing
    If Not m_cDragImage Is Nothing Then Set m_cDragImage = Nothing
    If Not m_cTransitionMask Is Nothing Then Set m_cTransitionMask = Nothing
    If Not m_cGridBuffer Is Nothing Then Set m_cGridBuffer = Nothing
    If Not m_pISelectorBar Is Nothing Then Set m_pISelectorBar = Nothing
    If Not m_cSelectorBar Is Nothing Then Set m_cSelectorBar = Nothing

End Sub

Private Function DestroyImlState() As Boolean
'*/ destroy header image list

    If Not (m_lImlStateHndl = 0) Then
        If ImageList_Destroy(m_lImlStateHndl) Then
            DestroyImlState = True
            m_lImlStateHndl = 0
        End If
    End If

End Function

Private Function DestroyImlDrag() As Boolean
'*/ destroy header image list

    If Not (m_lDragImgIml = 0) Then
        If ImageList_Destroy(m_lDragImgIml) Then
            DestroyImlDrag = True
            m_lDragImgIml = 0
        End If
    End If

End Function

Private Sub DestroyClasses()

Dim lUb As Long
Dim lCt As Long

    If ArrayCheck(m_cCellHeader) Then
        lUb = UBound(m_cCellHeader)
        For lCt = 0 To lUb
            Set m_cCellHeader(lCt) = Nothing
        Next lCt
    End If
    If ArrayCheck(m_cSubCell) Then
        lUb = UBound(m_cSubCell)
        For lCt = 0 To lUb
            Set m_cSubCell(lCt) = Nothing
        Next lCt
    End If
    If ArrayCheck(m_cControlDc) Then
        lUb = UBound(m_cControlDc)
        For lCt = 0 To lUb
            Set m_cControlDc(lCt) = Nothing
        Next lCt
    End If
    If Not (m_cEditBox Is Nothing) Then Set m_cEditBox = Nothing
    If Not m_cHGridSubclass Is Nothing Then Set m_cHGridSubclass = Nothing
    If Not m_cFilterMenu Is Nothing Then Set m_cFilterMenu = Nothing
    If Not m_cEditor Is Nothing Then Set m_cEditor = Nothing
    If Not m_cSkinHeader Is Nothing Then Set m_cSkinHeader = Nothing
    If Not m_cHeaderIcons Is Nothing Then Set m_cHeaderIcons = Nothing
    If Not m_cCellIcons Is Nothing Then Set m_cCellIcons = Nothing
    If Not m_cTreeIcons Is Nothing Then Set m_cTreeIcons = Nothing
    If Not m_cDragImage Is Nothing Then Set m_cDragImage = Nothing
    If Not m_cSkinScrollBars Is Nothing Then Set m_cSkinScrollBars = Nothing
    If Not m_cCellTips Is Nothing Then Set m_cCellTips = Nothing
    If Not m_cRender Is Nothing Then Set m_cRender = Nothing
    If Not m_cSizerDc Is Nothing Then Set m_cSizerDc = Nothing
    If Not OwnerDrawImpl Is Nothing Then Set OwnerDrawImpl = Nothing
    If Not c_ColumnTags Is Nothing Then Set c_ColumnTags = Nothing
    If Not c_PtrMem Is Nothing Then Set c_PtrMem = Nothing
    Erase m_tDividerRect

End Sub

Private Sub DestroyItems()
'/* destroy item classes

Dim lCt As Long
Dim lUb As Long

    If ArrayCheck(m_cGridItem) Then
        lUb = UBound(m_cGridItem)
        For lCt = 0 To lUb
            Set m_cGridItem(lCt) = Nothing
        Next lCt
    End If
    DestroyFonts

End Sub

Private Function DestroyList() As Boolean
'/* cleanup

    DestroyImlState
    DestroyImlDrag
    If Not (m_lHGHwnd = 0) Then
        If DestroyWindow(m_lHGHwnd) Then
            DestroyList = True
            m_lHGHwnd = 0
        End If
    End If
    DestroyArrays

End Function


'**********************************************************************
'*                              FIN
'**********************************************************************

'/~ It's been fun.. see you on the darkside (.net)
'/~ John
