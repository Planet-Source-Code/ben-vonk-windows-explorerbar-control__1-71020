A simple explanation of the ExplorerBar
Events, Properties, Functions and Subs

Some notes:
-----------
Groups:
 The groups will be automatically arranged
 If a Special group exist it will always be set on top
 If a Detail group exist it will always be set on bottom

Container:
 Only 1 group with a container can be set at a time
 It can be set in a Normal group or Special group, normaly
 the Container will be set in a Special group
 For CommandButtons in a container, set Style to Graphic
 for changing the backcolor of the button

Keys:
 When clicked a group or item, after that you can also use the
 cursor and tab keys to move up and down in the groups
 With the Space or Return key the group or item can be hit
 After a click of the mouse the Focus rectangle will be removed


Events:
-------
 Collapse(Group As Integer)
   Will be raised when a group collapse
   Returns:
   Group - number of group

 Expand(Group As Integer)
   Will be raised when a group expand
   Returns:
   Group - number of group

 ErrorOpenFile(Group As Integer, Item As Integer, File As String, Error As Long)
   Will be raised when a error occured while execute the file
   Returns:
   Group - number of group
   Item  - number of item
   File  - name of file
   Error - number of error

 GroupClick(Group As Integer, WindowState As WindowStates)
   Will be raised when a group is clicked
   Returns:
   Group       - number of group
   WindowState - windowstate of group
   0 = Expanded
   1 = Collapsed
   2 = Fixed

 ItemClick(Group As Integer, Item As Integer)
   Will be raised when a item is clicked
   Returns:
   Group - number of group
   Item  - number of item

 ItemOpenFile(Group As Integer, Item As Integer, File as String)
   Will be raised when a file is executed
   Returns:
   Group - number of group
   Item  - number of item
   File  - name of file

 MouseDown(Group As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Will be raised when the mouse button is pressed down
   Returns:
   Group  - number of group
   Button - button pressed
   Shift  - shift key is pressed
   X      - mouse X position
   Y      - mouse Y position

 MouseHover(Group As Integer, Item As Integer, FullTextShowed As Boolean)
   Will be raised when the mouse is hover
   Returns:
   Group          - number of group (-1 no group hovered, -2 mouse hovered scrollbar)
   Item           - number of item  (-1 no item hovered)
   FullTextShowed - ttrue if full text of Group or Item is showed

 MouseMove(Group As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Will be raised when the mouse is moved in a group
   Returns:
   Group  - number of group
   Button - button pressed
   Shift  - shift key is pressed
   X      - mouse X position
   Y      - mouse Y position

 MouseOut(Group As Integer, Item As Integer)
   Will be raised when the mouse is out
   Returns:
   Group - number of group (-1 out ExplorerBar, -2 out ScrollBar)
   Item  - number of item  (-1 no item)

 MouseUp(Group As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Will be raised when the mouse button is pressed up
   Returns:
   Group  - number of group
   Button - button pressed
   Shift  - shift key is pressed
   X      - mouse X position
   Y      - mouse Y position



Properties:
-----------
 Animation
   To set the animation for open and close groups

 BackColor
   To set the backcolor of the ExplorerBar

 BorderColor
   To set the bordercolor of the ExplorerBar
   
 DetailsForeColor
   To set the forecolor of the Details group

 DetailGroupButton
   To set a button on the Details group

 Font
   To set the Font of the ExplorerBar, groups and items

 GradientBackColor
   To set the gradientcolor of the ExplorerBar

 GradientNormalHeaderBackColor
   To set the gradientcolor of the NormalHeader group

 GradientNormalItemBackColor
   To set the gradientcolor of the NormalItem window

 GradientSpecialHeaderBackColor
   To set the gradientcolor of the SpecialHeader group

 GradientSpecialItemBackColor
   To set the gradientcolor of the SpecialItem window

 GradientStyle
   To set the style of the gradientcolor
   Values:
   - Top to Bottom
   - Bottom to Top
   - Left to Right
   - Right to Left

 HeaderHeight
   To set the height of the group header
   Values:
   - Low
   - High

 Locked
   Set on True to improve speed while add or delete Groups or items
   After the add or delete action, do not forget to set this property on false!

 NormalArrowDownColor
   To set the color when the arrow is down for the NormalHeader group
   
 NormalArrowHoverColor
   To set the hover when the arrow is hovered for the NormalHeader group

 NormalArrowUpColor
   To set the color when the arrow is up for the NormalHeader group

 NormalButtonBackColor
   To set the button backcolor for the NormalHeader group

 NormalButtonDownColor
   To set the color when the button is down for the NormalHeader group

 NormalButtonHoverColor
   To set the color when the button is hovered for the NormalHeader group

 NormalButtonUpColor
   To set the color when the button is up for the NormalHeader group

 NormalButtonPictureDown
   To set the picture when the button is down for the NormalHeader group

 NormalButtonPictureHover
   To set the picture when the button is hovered for the NormalHeader group

 NormalButtonPictureUp
   To set the picture when the button is up for the NormalHeader group

 NormalHeaderBackColor
   To set the backcolor for the NormalHeader group

 NormalHeaderForeColor
   To set the forecolor for the NormalHeader group

 NormalHeaderHoverColor
   To set the hovercolor for the NormalHeader group

 NormalItemBackColor
   To set the backcolor for the NormalItems group

 NormalItemBorderColor
   To set the bordercolor for the NormalItems group

 NormalItemForeColor
   To set the forecolor for the NormalItems group

 NormalItemHoverColor
   To set the hovercolor for the NormalItems group

 OpenOneGroupOnly
   If set on True only one group can be open at a time
   If the Deatail group not have a button it will allways open and not count as a group

 ShowBorder
   Shows the ExplorerBar border

 SoundGroupClicked
   Produce an sound if a group is clicked or entered

 SoundItemClicked
   Produce an sound if a item is clicked or entered

 SpecialArrowDownColor
   To set the color when the arrow is down for the SpecialHeader group

 SpecialArrowHoverColor
   To set the hover when the arrow is hovered for the SpecialHeader group

 SpecialArrowUpColor
   To set the color when the arrow is up for the SpecialHeader group

 SpecialButtonBackColor
   To set the button backcolor for the SpecialHeader group

 SpecialButtonDownColor
   To set the color when the button is down for the SpecialHeader group

 SpecialButtonHoverColor
   To set the color when the button is hovered for the SpecialHeader group

 SpecialButtonUpColor
   To set the color when the button is up for the SpecialHeader group

 SpecialButtonPictureDown
   To set the picture when the button is down for the SpecialHeader group

 SpecialButtonPictureHover
   To set the picture when the button is hovered for the SpecialHeader group

 SpecialButtonPictureUp
   To set the picture when the button is up for the SpecialHeader group

 SpecialHeaderBackColor
   To set the backcolor for the SpecialHeader group

 SpecialHeaderForeColor
   To set the forecolor for the SpecialHeader group

 SpecialHeaderHoverColor
   To set the hovercolor for the SpecialHeader group

 SpecialItemBackColor
   To set the backcolor for the SpecialItems group

 SpecialItemBorderColor
   To set the bordercolor for the SpecialItems group

 SpecialItemForeColor
   To set the forecolor for the SpecialItems group

 SpecialItemHoverColor
   To set the hovercolor for the SpecialItems group

 UseAlphaBlend() As Boolean
   To get fade in and fade out effect when Animation is set True

 UseTheme
   To use the type of Theme
   Values:
   - Windows
   - Enhanced
   - User

 UseUserForeColors
   To force to use the user defined colors
   Values:
   - True
   - False


Functions:
----------
 AddDetailGroup
   To add a Detail group, returns the number of groups
   Only 1 Detail group is possible
   the Maximum of total groups is 20
   Returns:
   -1  If a Details group is already add
   -20 If number of total groups is 20
   Parameters:
   - Title
   - TitleBold     (Optional)
     Default is True
   - WindowState   (Optional)
     Has effect only when the Detail group has a button
   - Icon          (Optional)
   - DetailPicture (Optional)
   - DetailTitle   (Optional)
   - DetailCaption (Optional)
   - ToolTipText   (Optional)
   - Tag           (Optional)

 AddItem
   To add a item to the given group, returns the number of items in the group
   The maximum of items per group is 20
   Returns:
   -1  If group not exist
   -2  If group has a container
   -3  If group is a Detail group
   -20 if number of items is 20
   Parameters:
   - Group
   - Caption
   - CaptionBold (Optional)
   - Icon        (Optional)
   - TextOnly    (Optional)
     If is True, the item is not a link
   - OpenFile    (Optional)
     When the item is clicked or pressed the given file will be executed
   - Tag         (Optional)
   - ToolTipText (Optional)

 AddNormalGroup
   To add a Normal group, returns the number of groups
   the Maximum of total groups is 20
   Returns:
   -20 If number of total groups is 20
   Parameters:
   - Title
   - TitleBold       (Optional)
     Default is True
   - WindowState     (Optional)
   - Icon            (Optional)
   - ItemsBackground (Optional)
   - ToolTipText     (Optional)
   - Tag             (Optional)
   - Container       (Optional)
     Only 1 group with a container can be set at a time
     Normaly the Container will be set in a Special group
     For CommandButtons in a container, set Style to Graphic
     for changing the backcolor of the button

 AddSpecialGroup
   To add a Special group, returns the number of groups
   the Maximum of total groups is 20
   Returns:
   -1  If a Special group is already add
   -20 If number of total groups is 20
   Parameters:
   - Title
   - TitleBold       (Optional)
     Default is True
   - WindowState     (Optional)
   - Icon            (Optional)
   - ItemsBackground (Optional)
   - ToolTipText     (Optional)
   - Tag             (Optional)
   - Container       (Optional)
     Only 1 group with a container can be set at a time
     For CommandButtons in a container, set Style to Graphic
     for changing the backcolor of the button

 DeleteGroup
   To delete a group, returns the number of groups
   returns the number of total groups
   Parameter:
   - Group

 DeleteItem
   To delete a item form the given group
   Returns the number of items in a group
   Parameters:
   - Group
   - Item

 FullTextShowed
   To check if the full Title or Caption is showed
   Returns True if is set
   Parameters:
   - Group
   - Item (Optional)

 GetDetailCaption
   Returns the group DetailCaption
   Parameter:
   - Group

 GetDetailPicture
   Returns the group DetailPicture
   Parameter:
   - Group

 GetDetailTitle
   Returns the group DetailTitle
   Parameter:
   - Group

 GetGroupContainer
   Returns the Container in a group
   Parameter:
   - Group

 GetGroupIcon
   Returns the group Icon
   Parameter:
   - Group

 GetGroupsCount
   Returns the number of groups

 GetGroupState
   Returns the group State
   1 = Normal
   2 = Hot
   3 = Pressed
   Parameter:
   - Group

 GetGroupTag
   Returns the group Tag
   Parameter:
   - Group

 GetGroupTitle
   Returns the group Title
   Parameter:
   - Group

 GetGroupTitleBold
   Returns the group TitleBold
   Parameter:
   - Group

 GetGroupToolTipText
   Returns the group ToolTipText
   Parameter:
   - Group

 GetGroupWindowState
   Returns the WindowState of a group
   0 = Expanded
   1 = Collapsed
   2 = Fixed
   Parameter:
   - Group

 GetItemCaption
   Returns the item Caption in a group
   Parameters:
   - Group
   - Item
   - StripTab (Optional)
     To remove the tabs from the caption in the result string

 GetItemCaptionBold
   Returns the item FontBold in a group
   Parameters:
   - Group
   - Item

 GetItemIcon
   Returns the item Icon in a group
   Parameters:
   - Group
   - Item

 GetItemOpenFile
   Returns the item OpenFile in a group
   Parameters:
   - Group
   - Item

 GetItemsBackgroundPicture
   Returns the item BackgroundPicture in a group
   Parameter:
   - Group

 GetItemsCount
   Returns the number of items in a group
   If group does not exist -1 will be returned
   Parameter:
   - Group

 GetItemTag
   Returns the item Tag in a group
   Parameters:
   - Group
   - Item

 GetItemTextOnly
   Returns the item TextOnly in a group
   Parameters:
   - Group
   - Item

 GetItemToolTipText
   Returns the item ToolTipText in a group
   Parameters:
   - Group
   - Item

 GetThemeMap
   Returns the Map of the theme

 GetThemeName
   Returns the Name of the theme

 GetThemeColorMap
   Returns the Map of the theme color

 GetThemeColorName
   Returns the Name of the theme color

 GetVersion
   Returns the version of the ExplorerBar

 hWnd
   Returns the hWnd of the ExplorerBar

 SetDetailCaption
   Sets the Detail group DetailCaption
   Returns True if is set
   Parameters:
   - Group
   - NewCaption

 SetDetailPicture
   Sets the Detail group DetailPicture
   Returns True if is set
   Parameters:
   - Group
   - NewPicture (Optional)

 SetDetailTitle
   Sets the Detail group DetailTitle
   Returns True if is set
   Parameters:
   - Group
   - NewTitle

 SetGroupContainer
   Sets the group Container
   Returns True if is set
   Only 1 container for the ExplorerBar can be set
   otherwise the return value will also be False
   Parameters:
   - Group
   - NewContainer (Optional)

 SetGroupIcon
   Sets the group Icon
   Returns True if is set
   Parameters:
   - Group
   - NewIcon (Optional)

 SetGroupTag
   Sets the group Tag
   Returns True if is set
   Parameters:
   - Group
   - NewTag (Optional)

 SetGroupTitle
   Sets the group Title
   Returns True if is set
   Parameters:
   - Group
   - NewTitle

 SetGroupTitleBold
   Sets the group Title FontBold
   Returns True if is set
   Parameters:
   - Group
   - NewTitleBold (Optional)

 SetGroupToolTipText
   Sets the group ToolTipText
   Returns True if is set
   Parameters:
   - Group
   - NewToolTipText (Optional)

 SetGroupWindowState
   Sets the group WindowState
   Returns True if is set
   Parameters:
   - Group
   - NewWindowState (Optional)
     0 = Expanded
     1 = Collapsed
     2 = Fixed

 SetItemCaption
   Sets the item Caption in a group
   Returns True if is set
   Parameters:
   - Group
   - Item
   - NewCaption

 SetItemCaptionBold
   Sets the item Caption FontBold in a group
   Returns True if is set
   Parameters:
   - Group
   - Item
   - NewCaptionBold (Optional)

 SetItemIcon
   Sets the item Icon in a group
   Returns True if is set
   Parameters:
   - Group
   - Item
   - NewIcon (Optional)

 SetItemOpenFile
   Sets the item OpenFile in a group
   Returns True if is set
   Parameters:
   - Group
   - Item
   - NewOpenFile (Optional)
     When the item is clicked or pressed the given file will be executed

 SetItemsBackgroundPicture
   Sets the item BackgroundPicture in a group
   Returns True if is set
   Parameters:
   - Group
   - Item
   - NewBackgroundPicture (Optional)

 SetItemTag
   Sets the item Tag in a group
   Returns True if is set
   Parameters:
   - Group
   - Item
   - NewTag (Optional)

 SetItemTextOnly
   Sets the item TextOnly in a group
   Returns True if is set
   Parameters:
   - Group
   - Item
   - NewTextOnly (Optional)
     If is True, the item is not a link

 SetItemToolTipText
   Sets the item ToolTipText in a group
   Returns True if is set
   Parameters:
   - Group
   - Item
   - NewToolTipText (Optional)


Subs:
-----
 DeleteAllGroupItems
   To delete all items in the specified group

 DeleteAllGroups
   To delete all groups from the ExplorerBar

 Refresh
   To refresh the ExplorerBar
