<?xml version="1.0" encoding="utf-8"?>
<styleLibrary>
  <annotation>
    <lastModified>2007-09-24T14:57:54</lastModified>
  </annotation>
  <styleSets defaultStyleSet="Default">
    <styleSet name="Default" buttonStyle="FlatBorderless" viewStyle="Standard" useOsThemes="False" useFlatMode="True">
      <componentStyles>
        <componentStyle name="Inbox Form">
          <properties>
            <property name="BackColor" colorCategory="{Default}">Gainsboro</property>
          </properties>
        </componentStyle>
        <componentStyle name="Inbox MonthCalendar">
          <properties>
            <property name="TitleBackColor" colorCategory="{Default}">Black</property>
          </properties>
        </componentStyle>
        <componentStyle name="Inbox Panel">
          <properties>
            <property name="BackColor" colorCategory="{Default}">Gainsboro</property>
          </properties>
        </componentStyle>
        <componentStyle name="UltraButton" buttonStyle="FlatBorderless" useFlatMode="True" />
        <componentStyle name="UltraCalculatorDropDown" useFlatMode="True" />
        <componentStyle name="UltraCalendarCombo" useOsThemes="False" useFlatMode="True" />
        <componentStyle name="UltraCheckEditor" useOsThemes="False" useFlatMode="True" />
        <componentStyle name="UltraColorPicker" useFlatMode="True" />
        <componentStyle name="UltraCombo" useOsThemes="False" useFlatMode="True" />
        <componentStyle name="UltraComboEditor" useFlatMode="True" />
        <componentStyle name="UltraCurrencyEditor" useFlatMode="True" />
        <componentStyle name="UltraDateTimeEditor" useFlatMode="True" />
        <componentStyle name="UltraDayView" useOsThemes="True" useFlatMode="True" />
        <componentStyle name="UltraDockManager" buttonStyle="FlatBorderless">
          <properties>
            <property name="GroupPaneTabStyle" colorCategory="{Default}">Flat</property>
            <property name="UnpinnedTabStyle" colorCategory="{Default}">Flat</property>
          </properties>
        </componentStyle>
        <componentStyle name="UltraDropDown" useOsThemes="False" useFlatMode="True" />
        <componentStyle name="UltraFontNameEditor" useFlatMode="True" />
        <componentStyle name="UltraGrid" useOsThemes="False" useFlatMode="True" />
        <componentStyle name="UltraMaskedEdit" useFlatMode="True" />
        <componentStyle name="UltraNumericEditor" useFlatMode="True" />
        <componentStyle name="UltraStatusBar" useFlatMode="True" />
        <componentStyle name="UltraTabbedMdiManager">
          <properties>
            <property name="TabStyle" colorCategory="{Default}">Flat</property>
          </properties>
        </componentStyle>
        <componentStyle name="UltraTabControl" useOsThemes="False" useFlatMode="True">
          <properties>
            <property name="Style" colorCategory="{Default}">Flat</property>
          </properties>
        </componentStyle>
        <componentStyle name="UltraTabStripControl" useOsThemes="False" useFlatMode="True">
          <properties>
            <property name="Style" colorCategory="{Default}">Flat</property>
          </properties>
        </componentStyle>
        <componentStyle name="UltraTextEditor" useFlatMode="True" />
        <componentStyle name="UltraTimeZoneEditor" useFlatMode="True" />
        <componentStyle name="UltraToolbarsManager" useOsThemes="False" useFlatMode="True" />
        <componentStyle name="UltraTree" buttonStyle="FlatBorderless" headerStyle="Standard" useFlatMode="True" />
        <componentStyle name="UltraWeekView" useOsThemes="False" useFlatMode="True" />
      </componentStyles>
      <styles>
        <style role="AutoRepeatEditorButton">
          <states>
            <state name="Normal">
              <resources>
                <name>sharedButtonRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="Base">
          <states>
            <state name="Normal" foreColor="30, 30, 30" fontName="Tahoma" textVAlign="Middle" fontSize="8" themedElementAlpha="Transparent" />
            <state name="HotTracked" fontName="Tahoma" />
          </states>
        </style>
        <style role="Button" buttonStyle="FlatBorderless">
          <states>
            <state name="Normal" backColor="Transparent" foreColor="White" borderColor="Transparent" textHAlign="Center" backGradientStyle="None">
              <resources>
                <name>sharedButtonRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="Selected" backColor="Transparent" backGradientStyle="None" />
            <state name="HotTracked" foreColor="Silver" borderColor="Transparent" />
            <state name="Pressed" backColor="Transparent" borderColor="Transparent" backGradientStyle="None" />
            <state name="Active" backColor="Transparent" backGradientStyle="None" />
            <state name="Checked" backColor="Transparent" backGradientStyle="None" />
          </states>
        </style>
        <style role="CalendarComboControlArea">
          <states>
            <state name="Normal" borderColor="DarkGray" />
          </states>
        </style>
        <style role="CalendarComboDateButton">
          <states>
            <state name="Normal" backColor="Transparent" foreColor="White" borderColor="Transparent" imageBackgroundStyle="Stretched" backGradientStyle="None" imageBackgroundStretchMargins="3, 7, 4, 5" />
          </states>
        </style>
        <style role="CalendarComboDateButtonArea">
          <states>
            <state name="Normal" backColor="237, 237, 237" borderColor="85, 85, 85" backGradientStyle="None" />
          </states>
        </style>
        <style role="CalendarComboDropDown">
          <states>
            <state name="Normal" backColor="237, 237, 237" borderColor="Black" backGradientStyle="None" />
          </states>
        </style>
        <style role="CheckEditor">
          <states>
            <state name="Normal" backColor="Transparent" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
            <state name="HotTracked" backColor="White" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
            <state name="Checked" backColor="Transparent" backGradientStyle="None" />
            <state name="Unchecked" backColor="Transparent" backGradientStyle="None" />
          </states>
        </style>
        <style role="DayViewAllDayEventArea">
          <states>
            <state name="Normal" backColor="237, 237, 237" borderColor="Transparent" backGradientStyle="None" />
            <state name="Selected" backColor="White" backColor2="237, 237, 237" backGradientStyle="Vertical" />
          </states>
        </style>
        <style role="DayViewTimeSlotDescriptor">
          <states>
            <state name="Normal" backColor="237, 237, 237" fontBold="True" backGradientStyle="None" />
          </states>
        </style>
        <style role="DayViewTimeSlotNonWorkingHour">
          <states>
            <state name="Normal" backColor="White" backColor2="237, 237, 237" />
            <state name="Selected" backColor="White" borderColor="DimGray" backColor2="225, 225, 225" backGradientStyle="Vertical" />
          </states>
        </style>
        <style role="DayViewTimeSlotWorkingHour">
          <states>
            <state name="Normal" backColor="Gainsboro" backGradientStyle="None" />
            <state name="Selected" backColor="White" backColor2="237, 237, 237" backGradientStyle="Vertical" />
          </states>
        </style>
        <style role="DesktopAlertCloseButton">
          <states>
            <state name="Normal" foreColor="Black">
              <resources>
                <name>spacer</name>
              </resources>
            </state>
            <state name="HotTracked" backColor="Transparent" foreColor="White" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None" borderColor2="Transparent">
              <resources>
                <name>scrollBarArrowDownRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="DesktopAlertControlArea">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="DesktopAlertDropDownButton">
          <states>
            <state name="Normal" foreColor="Black">
              <resources>
                <name>spacer</name>
              </resources>
            </state>
            <state name="HotTracked" foreColor="White">
              <resources>
                <name>scrollBarThumbHorizontalRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="DesktopAlertGripArea">
          <states>
            <state name="Normal">
              <resources>
                <name>headerRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="DesktopAlertLink">
          <states>
            <state name="Normal" foreColor="Black" textVAlign="Top" />
            <state name="HotTracked" foreColor="Black" />
          </states>
        </style>
        <style role="DesktopAlertPinButton">
          <states>
            <state name="Normal" foreColor="Black">
              <resources>
                <name>spacer</name>
              </resources>
            </state>
            <state name="HotTracked">
              <resources>
                <name>sharedButtonRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="DockAreaSplitterBarVertical">
          <states>
            <state name="Normal" backColor="Transparent" borderColor="Transparent" backGradientStyle="None" />
          </states>
        </style>
        <style role="DockControlPane">
          <states>
            <state name="Normal" backColor="237, 237, 237" borderColor="DarkGray" backGradientStyle="None" />
          </states>
        </style>
        <style role="DockControlPaneContentArea">
          <states>
            <state name="Normal" backColor="237, 237, 237" borderColor="DarkGray" backGradientStyle="None" />
          </states>
        </style>
        <style role="DockFloatingWindowCaptionHorizontal">
          <states>
            <state name="Normal" foreColor="210, 210, 210" borderColor="Transparent" />
            <state name="Active" foreColor="White" />
          </states>
        </style>
        <style role="DockPaneCaption">
          <states>
            <state name="Normal">
              <resources>
                <name>scheduleDayOfWeekHeaderRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="DockPaneCaptionHorizontal">
          <states>
            <state name="Normal" foreColor="195, 195, 195">
              <resources>
                <name>headerRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="Active" foreColor="White" />
          </states>
        </style>
        <style role="DockSlidingGroupHeaderHorizontal">
          <states>
            <state name="Normal">
              <resources>
                <name>scheduleDayOfWeekHeaderRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="DropDownButton" buttonStyle="FlatBorderless">
          <states>
            <state name="Normal" borderColor="Transparent">
              <resources>
                <name>dropDownButtonRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="HotTracked" backColor="Transparent" borderColor="Transparent" backGradientStyle="None" />
            <state name="Pressed" backColor="Transparent" borderColor="Transparent" backGradientStyle="None" />
          </states>
        </style>
        <style role="DropDownControlArea">
          <states>
            <state name="Normal" backColor="237, 237, 237" borderColor="DarkGray" backGradientStyle="None" />
          </states>
        </style>
        <style role="ExplorerBarControlArea">
          <states>
            <state name="Normal">
              <resources>
                <name>explorerBarControlAreaRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ExplorerBarGroupHeader">
          <states>
            <state name="Normal" foreColor="White" textVAlign="Middle">
              <resources>
                <name>headerRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="HotTracked">
              <resources>
                <name>headerRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ExplorerBarGroupItemAreaInner">
          <states>
            <state name="Normal">
              <resources>
                <name>explorerBarGroupItemAreaInnerRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ExplorerBarGroupItemAreaOuter">
          <states>
            <state name="Normal" backColor="Black" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ExplorerBarItem">
          <states>
            <state name="Normal">
              <resources>
                <name>explorerBarItemRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ExplorerBarNavigationOverflowButtonArea">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ExplorerBarNavigationSplitter">
          <states>
            <state name="Normal" backColor="DarkGray" borderColor="Gray" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="GridAddNewBox">
          <states>
            <state name="Normal" borderColor="DarkGray" fontBold="True">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="GridBandHeader">
          <states>
            <state name="Normal" backColor="Transparent" foreColor="White" borderColor="Transparent" fontName="Tahoma" textVAlign="Top" imageBackgroundStyle="Stretched" fontBold="True" fontSize="8" backGradientStyle="None" imageBackgroundStretchMargins="3, 7, 4, 5">
              <resources>
                <name>headerRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="HotTracked" borderColor="Transparent" textVAlign="Top" />
          </states>
        </style>
        <style role="GridCaption">
          <states>
            <state name="Normal" backColor="Transparent" foreColor="240, 240, 240" borderColor="Transparent" fontName="Tahoma" textVAlign="Top" imageBackgroundStyle="Stretched" fontBold="True" fontSize="8" backGradientStyle="None" imageBackgroundStretchMargins="3, 7, 4, 5">
              <resources>
                <name>headerRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="GridCardArea">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="GridCardCaption">
          <states>
            <state name="Normal" backColor="240, 240, 240" imageBackgroundStyle="Stretched" backGradientStyle="None" imageBackgroundStretchMargins="3, 7, 4, 5">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAVgEAAAKJUE5HDQoaCgAAAA1JSERSAAAALAAAABIIBgAAACD0PNwAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAv0lEQVRIS+3Wuw3EIBAEUOgJIZBFGVRDSDO0QgAFmA6I6GCtIUD23Tk8aS05GOKn5TdijEG9d/Lek5SShBCsAhNsMMIqsOz7PpFaa5aBDcYFLqVMsLWWZWCDcYFzzhPsnGMZ2GD8Am/bRhxzCzbGEMfcgjlfup9HgtuTdva84H/vzjvhd8IfneZyJGqtpJRiVXrOOwYbjOuna61RSolijBRCYBWYYINxgVHd0IZQMDB6ToEJtku9hPwpmX34STkAAXvpWlyqhaEAAAAASUVORK5CYIILAAAAAAAAAAA=</imageBackground>
            </state>
            <state name="Selected" backColor="240, 240, 240" imageBackgroundStyle="Stretched" backGradientStyle="None" />
          </states>
        </style>
        <style role="GridCell">
          <states>
            <state name="Normal" backColor="Transparent" foreColor="45, 45, 45" borderColor="Transparent" backGradientStyle="None" borderColor3DBase="Transparent" />
            <state name="Selected" backColor="Transparent" foreColor="30, 30, 30" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None" imageBackgroundStretchMargins="1, 1, 1, 1" />
            <state name="HotTracked">
              <resources>
                <name>gridCellRubberBlack_HotTracked</name>
              </resources>
            </state>
            <state name="Active" backColor="Transparent" foreColor="White" borderColor="Transparent" fontName="Tahoma" fontSize="8" backGradientStyle="None" backHatchStyle="None" />
            <state name="EditMode" backColor="237, 237, 237" foreColor="DimGray" imageBackgroundStyle="Stretched" fontUnderline="True" backGradientStyle="None" imageBackgroundStretchMargins="2, 2, 2, 2">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAACgEAAAKJUE5HDQoaCgAAAA1JSERSAAAAKAAAAA4IBgAAAF0LXkYAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAc0lEQVRIS2OcOXNmAgMDQz8QCwDxfyAeLOAD0CFFDEAH/gVih////zMMJgx0UwjIbSAH/hlMDkN2y6gDKY2Z0RAcDUFKQ4BS/aNpcFiHIDB6WUCVCKyqcwdyOCj1MbX0g9wCxBGwNBgPZLwBuXaQYZCb4gD2LAGKQyV8jQAAAABJRU5ErkJgggsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
            </state>
          </states>
        </style>
        <style role="GridColumnHeader">
          <states>
            <state name="Normal" backColor="Transparent" foreColor="White" borderColor="Transparent" fontName="Tahoma" textHAlign="Left" textVAlign="Middle" imageBackgroundStyle="Stretched" fontBold="True" fontSize="8" backGradientStyle="None" imageBackgroundStretchMargins="3, 7, 4, 5">
              <resources>
                <name>headerRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="HotTracked" backColor="Transparent" foreColor="White" borderColor="Transparent" fontName="Tahoma" textVAlign="Middle" imageBackgroundStyle="Stretched" fontBold="True" fontSize="8" backGradientStyle="None" imageBackgroundStretchMargins="3, 7, 4, 5">
              <cursor>Hand</cursor>
              <resources>
                <name>headerRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="GridControlArea">
          <states>
            <state name="Normal" backColor="White" borderColor="Transparent" fontName="Tahoma" fontSize="8" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="GridGroupByBox">
          <states>
            <state name="Normal" borderColor="DarkGray">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="GridGroupByBoxPrompt">
          <states>
            <state name="Normal" backColor="30, 30, 30" foreColor="White" borderColor="DarkGray" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="GridGroupByRow">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="GridRow">
          <states>
            <state name="Normal" backColor="Transparent" borderColor="DarkGray" fontName="Tahoma" fontSize="8" backGradientStyle="None" backHatchStyle="None" />
            <state name="Selected">
              <resources>
                <name>gridRowRubberBlack_Selected</name>
              </resources>
            </state>
            <state name="Active">
              <resources>
                <name>gridRowRubberBlack_Active</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="GridRowScrollRegionSplitBox">
          <states>
            <state name="Normal">
              <resources>
                <name>sharedButtonRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="GridRowScrollRegionSplitterBar">
          <states>
            <state name="Normal">
              <resources>
                <name>sharedButtonRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="GridRowSelector">
          <states>
            <state name="Normal" backColor="Transparent" borderColor="Transparent" backGradientStyle="None" />
          </states>
        </style>
        <style role="GridRowSelectorHeader">
          <states>
            <state name="Normal" backColor="Transparent" borderColor="Transparent" backGradientStyle="None" borderColor3DBase="Transparent" />
          </states>
        </style>
        <style role="GroupPaneTabItemAreaHorizontal">
          <states>
            <state name="Normal" backColor="237, 237, 237" backGradientStyle="None" />
          </states>
        </style>
        <style role="GroupPaneTabItemAreaHorizontalBottom">
          <states>
            <state name="Normal" backColor="237, 237, 237" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="GroupPaneTabItemAreaVertical">
          <states>
            <state name="Normal" backColor="237, 237, 237" backGradientStyle="None" />
          </states>
        </style>
        <style role="GroupPaneTabItemHorizontal">
          <states>
            <state name="Normal" foreColor="210, 210, 210" imageBackgroundStyle="Stretched" imageBackgroundStretchMargins="7, 6, 7, 10">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAArAIAAAKJUE5HDQoaCgAAAA1JSERSAAAAVQAAABMIBgAAAOH0QSsAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAACFUlEQVRYR+3Yu64pURgH8LFVCkI0Cl5AUCjI2TqNQlwKtxARl7hlE5ctQUJQi8QDSFQukXgFhVIi6n3eYFdegP+Zb9mcnJPsGOc0irWSf9bMLMXkl29mfUP28vKC2Wwm6PV6wWAwCEqlUhCvCXxIFzidTsLxeBQ+Pz8Fh8MhECDG4zFWqxU2mw12ux0OhwPPAwb7/R7b7Rbr9RrkKchkMjQaDQyHQ0wmE8zncyyXSzbz3DdYLBagTKdTjEYjkCdDDYfDeHt7Q7vdRr/fZ+l0OjwSDLrdLgaDAbMql8u/UcX3ADweDyKRCOLxOKLRKAKBAI8Eg1AodDPz+XwXVHFA3KBgNpths9lgt9thtVphNBp5JBiYTKabm8ViIdALKkXc9aHVaqHRaKBQKK6LfP7yuTr9PcvlcqjVamanUqk46ndQj1znqHeq7hHM6285Kke9vOufPd9WqtPpRDAYRCqVQjabRSaTQSKR4JFgkEwmkcvlkE6nWb9/2/1dLhdisRhbLJVKKBaLDJfnvkE+n2dmhUKB9as3VLfbzaqSvqpqtRqbeaQZ0FcUmdFMVXtFPfv9fvbI0+L7+zvq9TqPRAPyajabzI2edBH1TM3/T6/Xi2q1ilarxfMfBl+V+kGorzqd7lypVNDr9Xj+0YD8yFH0/HH9J7YsHpyevX158vsjUHL8Y7yKZ1S6tPj0/eET3SN5kRv5sfELj1fgOV+yJvQAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==</imageBackground>
            </state>
            <state name="Selected" foreColor="White" fontBold="True" />
          </states>
        </style>
        <style role="GroupPaneTabItemHorizontalBottom">
          <states>
            <state name="Normal" borderColor="Transparent" />
          </states>
        </style>
        <style role="GroupPaneTabItemVertical">
          <states>
            <state name="Normal" backColor="Transparent" foreColor="210, 210, 210" borderColor="Transparent" imageBackgroundStyle="Stretched" backGradientStyle="None" imageBackgroundStretchMargins="6, 10, 6, 10">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAApAIAAAKJUE5HDQoaCgAAAA1JSERSAAAAEwAAAFUIBgAAAKtR370AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAACDUlEQVRYR+1ZuaoiQRTtQVFE3BBxCURFUDRQUBP3xA1E3DBwx0ARTF1Q8QuExiUV3h/MN/gFMx/hV5ic6SrnOfim3/DCCW7DCbq7+tTtrttVp879djwecbvdhI+43+/CF46fUpsDgDfedjweo16vI5PJIBgMwmazQalUQrolC61WC6/Xi2QyiUQiAbVazdp9lwgFYbfbYTqdotFoIJ1Ow+/3w2w2Q6/Xy8LhcCAWi6FarWI2m2G5XL53KgqiKPILg8EAhUIB4XCYR8cI5eB2u3mn3W4X6/Ual8sFzWbzQXg6nbDZbDAajVAsFhEKhWCxWGA0GmXhdDo5Wb/f588xssPh8CBjA7BYLNDpdJDNZvn30Gg0UCgUsmDRstdstVrPyK7X6ytZu93mPXo8HqhUqk8HgEUciUSITBoA+mYvaUKp8effpNSg1HisATTT0oIirdo0OdLk+ImmpdSg1KDUoE3FY1dHypGUIylHsiKefwHJA5IHJA9IHvyv8oB528zwHg6HyOVyCAQCMBgM3JKWg91uRzwe5/b1X972+Xx+GuWlUom77larFSaTSRYul4vXDphRzuz/F9d9v99jPp+j1+shn89z1/1fFj4jS6VSXIKtVitO+LsQIQrb7RaTyQS1Wo1XH3w+H49Ip9PJgr1mNBpFpVJBuVzmNYNnpeLdwv+q2PtQDvkhnQ9YyYPhFyTfbBuNW2o6AAAAAElFTkSuQmCCCwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA</imageBackground>
            </state>
            <state name="Selected" foreColor="White" fontBold="True" />
          </states>
        </style>
        <style role="GroupPaneTabItemVerticalRight">
          <states>
            <state name="Normal" borderColor="Transparent" />
          </states>
        </style>
        <style role="ListViewColumnHeader">
          <states>
            <state name="Normal" borderColor="Transparent">
              <resources>
                <name>headerRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ListViewControlArea">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ListViewGroupHeader">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="MainMenubarHorizontal">
          <states>
            <state name="Normal">
              <resources>
                <name>mainMenubarHorizontalRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="MaskLiteralChar">
          <states>
            <state name="Normal" backColor="Transparent" backGradientStyle="None">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="MenuCheckMark">
          <states>
            <state name="Normal" backColor="Transparent" foreColor="White" borderColor="Transparent" backGradientStyle="None" />
          </states>
        </style>
        <style role="MenuItem">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="HotTracked">
              <resources>
                <name>sharedButtonRubberBlack_HotTracked</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="MenuItemAddRemoveTool">
          <states>
            <state name="Normal" backColor="237, 237, 237" foreColor="30, 30, 30" borderColor="85, 85, 85" fontName="Tahoma" fontSize="8" backGradientStyle="None" />
            <state name="HotTracked">
              <resources>
                <name>sharedButtonRubberBlack_HotTracked</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="MenuItemComboBox">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="HotTracked">
              <resources>
                <name>sharedButtonRubberBlack_HotTracked</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="MenuItemProgressBar">
          <states>
            <state name="Normal" backColor="237, 237, 237" backGradientStyle="None" backHatchStyle="None">
              <resources>
                <name>sharedButtonRubberBlack_HotTracked</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="MenuItemResetToolbar">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="MenuItemTaskPane">
          <states>
            <state name="Normal" foreColor="White" />
            <state name="HotTracked" foreColor="Black" />
          </states>
        </style>
        <style role="MenuItemToolbarPicker">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="HotTracked">
              <resources>
                <name>sharedButtonRubberBlack_HotTracked</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="MenuSideStrip">
          <states>
            <state name="Normal" backColor="240, 240, 240" fontBold="True" backColor2="195, 195, 195" backGradientStyle="Horizontal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="MenuTearAwayStrip">
          <states>
            <state name="Normal" borderColor="237, 237, 237">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="MonthViewMultiControlArea">
          <states>
            <state name="Normal" backColor="237, 237, 237" foreColor="30, 30, 30" backGradientStyle="None" />
          </states>
        </style>
        <style role="MonthViewMultiScrollButton">
          <states>
            <state name="Normal" backColor="Transparent" borderColor="White" backGradientStyle="None" />
          </states>
        </style>
        <style role="MonthViewSingleControlArea">
          <states>
            <state name="Normal" borderColor="DarkGray" />
          </states>
        </style>
        <style role="NavigationBarActionButton">
          <states>
            <state name="Normal" backColor="Transparent" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="NavigationBarLocationButton">
          <states>
            <state name="Normal" foreColor="0, 0, 0">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAA+QAAAAKJUE5HDQoaCgAAAA1JSERSAAAAKAAAACgIBgAAAIz+uG0AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsRAAALEQF/ZF+RAAAAYklEQVRYR+3SsQ3AMAwDwXj/oRU4K3wKFWfAJQHixDMzz+p3C27+q8t9192sp+Af13HiqkiQYBWoeRskWAVq3gYJVoGat0GCVaDmbZBgFah5GyRYBWreBglWgZq3QYJVoOZfWTs5xeZ58m0AAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
            </state>
            <state name="HotTracked" foreColor="255, 255, 255">
              <resources>
                <name>headerRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="Pressed" backColor="Transparent" foreColor="255, 255, 255" backGradientStyle="None" backHatchStyle="None">
              <resources>
                <name>headerRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="NavigationBarLocationListItem">
          <states>
            <state name="HotTracked" foreColor="255, 255, 255">
              <resources>
                <name>headerRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="NavigationBarPreviousLocationsDropDownButton">
          <states>
            <state name="Normal" foreColor="0, 0, 0">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAA+QAAAAKJUE5HDQoaCgAAAA1JSERSAAAAKAAAACgIBgAAAIz+uG0AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsRAAALEQF/ZF+RAAAAYklEQVRYR+3SsQ3AMAwDwXj/oRU4K3wKFWfAJQHixDMzz+p3C27+q8t9192sp+Af13HiqkiQYBWoeRskWAVq3gYJVoGat0GCVaDmbZBgFah5GyRYBWreBglWgZq3QYJVoOZfWTs5xeZ58m0AAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
            </state>
          </states>
        </style>
        <style role="NavigationToolbarBackButton">
          <states>
            <state name="Normal" backColor="Transparent" backGradientStyle="None" backHatchStyle="None">
              <imageBackground>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AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA</imageBackground>
            </state>
            <state name="HotTracked">
              <imageBackground>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AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA</imageBackground>
            </state>
            <state name="Pressed">
              <imageBackground>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AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA</imageBackground>
            </state>
          </states>
        </style>
        <style role="NavigationToolbarButtonArea">
          <states>
            <state name="Normal" imageBackgroundStyle="Centered" />
          </states>
        </style>
        <style role="NavigationToolbarForwardButton">
          <states>
            <state name="Normal" backColor="Transparent" backGradientStyle="None" backHatchStyle="None">
              <imageBackground>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AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==</imageBackground>
            </state>
            <state name="HotTracked">
              <imageBackground>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AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA</imageBackground>
            </state>
            <state name="Pressed">
              <imageBackground>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AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA</imageBackground>
            </state>
          </states>
        </style>
        <style role="NavigationToolbarRecentHistoryButton">
          <states>
            <state name="Normal" imageHAlign="Right" imageVAlign="Middle">
              <image>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAHAEAAAKJUE5HDQoaCgAAAA1JSERSAAAADQAAAA4IBgAAAPR/ltIAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAA7BAAAOwQG4kWvtAAAAhUlEQVQoU5WR0Q2AIAxEdQYdyjn8Yxu+nYAZ/GIgd0COADmahhSSJrTXxzV0TyltywfQaiwDZTLhcub8yuFq4I7a0McJxJuABqI2gAyxQwPYsfcy5GKMj/f+4wghvNVdhyAySAAcVaiPB1AA0FTomHwENBVCEaL88gHQ9mRatqlptlzzAz9tFAhyHfWzUgAAAABJRU5ErkJgggsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</image>
            </state>
            <state name="HotTracked">
              <image>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAALQEAAAKJUE5HDQoaCgAAAA1JSERSAAAADQAAAA4IBgAAAPR/ltIAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAA7BAAAOwQG4kWvtAAAAlklEQVQoU5WSzQ2AIAyFdRs4MwmMgxuwhkNwdgjH4OQA2BJKSjFESF74ab++QthzztvyQGhVHWCM0aADdFbhWsuiDYKgY8kE0ew4WKDqIBPlvjkS5CcuBHtyI6gEQghXSunhijHeVPATkiAHMCahrj10lABAQ3vqx51U51Rf0E5AOzw5HQCEjrxVXDeHwWnlKy3/Oyz+An1mcvIJVCGUAAAAAElFTkSuQmCCCwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</image>
            </state>
            <state name="Pressed">
              <image>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAALQEAAAKJUE5HDQoaCgAAAA1JSERSAAAADQAAAA4IBgAAAPR/ltIAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAA7BAAAOwQG4kWvtAAAAlklEQVQoU5WSzQ2AIAyFdRs4MwmMgxuwhkNwdgjH4OQA2BJKSjFESF74ab++QthzztvyQGhVHWCM0aADdFbhWsuiDYKgY8kE0ew4WKDqIBPlvjkS5CcuBHtyI6gEQghXSunhijHeVPATkiAHMCahrj10lABAQ3vqx51U51Rf0E5AOzw5HQCEjrxVXDeHwWnlKy3/Oyz+An1mcvIJVCGUAAAAAElFTkSuQmCCCwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</image>
            </state>
          </states>
        </style>
        <style role="PopupMenu">
          <states>
            <state name="Normal" backColor="Transparent" backGradientStyle="None" />
          </states>
        </style>
        <style role="PrintPreviewPageArea">
          <states>
            <state name="Normal" backColor="237, 237, 237" backGradientStyle="None" />
          </states>
        </style>
        <style role="ProgressBarFill">
          <states>
            <state name="Normal">
              <resources>
                <name>progressBarFillRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="QuickCustomizeToolbar">
          <states>
            <state name="Normal" backColor="237, 237, 237" backGradientStyle="None" />
          </states>
        </style>
        <style role="ScheduleAppointment">
          <states>
            <state name="Normal" backColor="White" backColor2="237, 237, 237" backGradientStyle="BackwardDiagonal" />
            <state name="Selected" backColor2="237, 237, 237" />
          </states>
        </style>
        <style role="ScheduleCurrentDay">
          <states>
            <state name="Selected" backColor="237, 237, 237" backGradientStyle="None" />
          </states>
        </style>
        <style role="ScheduleCurrentDayHeader">
          <states>
            <state name="Normal" backColor="White" backColor2="237, 237, 237" />
          </states>
        </style>
        <style role="ScheduleDay">
          <states>
            <state name="Normal">
              <resources>
                <name>scheduleDayRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="Selected">
              <resources>
                <name>scheduleDayRubberBlack_Selected</name>
              </resources>
            </state>
            <state name="Active">
              <resources>
                <name>scheduleDayRubberBlack_Active</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ScheduleDayHeader" borderStyle="None">
          <states>
            <state name="Normal">
              <resources>
                <name>scheduleDayHeaderRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="Selected">
              <resources>
                <name>scheduleDayHeaderRubberBlack_Selected</name>
              </resources>
            </state>
            <state name="Active" borderColor="Transparent" textHAlign="Left" />
          </states>
        </style>
        <style role="ScheduleDayOfWeekHeader" borderStyle="None">
          <states>
            <state name="Normal" backColor="237, 237, 237" foreColor="Black" borderColor="Transparent" fontBold="True" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="ScheduleMonth">
          <states>
            <state name="Normal" fontName="Tahoma" fontSize="8" />
          </states>
        </style>
        <style role="ScheduleMonthHeader">
          <states>
            <state name="Normal" foreColor="White" borderColor="Transparent" fontName="Verdana" textHAlign="Center" textVAlign="Top" imageBackgroundStyle="Stretched" fontBold="True" fontSize="7" imageBackgroundStretchMargins="3, 7, 4, 5">
              <resources>
                <name>scheduleDayOfWeekHeaderRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ScheduleOwnerHeader">
          <states>
            <state name="Normal" backColor="Transparent" foreColor="White" borderColor="237, 237, 237" backGradientStyle="None">
              <resources>
                <name>headerRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ScheduleTrailingDay">
          <states>
            <state name="Normal" borderColor="Transparent" fontName="Tahoma" fontSize="8" />
          </states>
        </style>
        <style role="ScheduleWeek">
          <states>
            <state name="Normal" backColor="237, 237, 237" fontName="Tahoma" fontSize="8" backGradientStyle="None" />
          </states>
        </style>
        <style role="ScheduleWeekHeader">
          <states>
            <state name="Normal" backColor="237, 237, 237" borderColor="Transparent" backGradientStyle="None">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ScrollBarArrowHorizontal">
          <states>
            <state name="Normal">
              <resources>
                <name>scrollBarThumbHorizontalRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ScrollBarArrowVertical">
          <states>
            <state name="Normal" backColor="Transparent" backGradientStyle="None" backHatchStyle="None">
              <resources>
                <name>scrollBarThumbVerticalRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ScrollBarIntersection">
          <states>
            <state name="Normal" backColor="237, 237, 237" backGradientStyle="None" />
          </states>
        </style>
        <style role="ScrollBarThumbHorizontal">
          <states>
            <state name="Normal">
              <resources>
                <name>scrollBarThumbHorizontalRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ScrollBarThumbVertical">
          <states>
            <state name="Normal">
              <resources>
                <name>scrollBarThumbVerticalRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ScrollBarTrack">
          <states>
            <state name="Normal" backColor="237, 237, 237" backGradientStyle="None" />
          </states>
        </style>
        <style role="ScrollBarTrackHorizontal">
          <states>
            <state name="Normal" backColor="White" backColor2="237, 237, 237" backGradientStyle="Vertical" />
          </states>
        </style>
        <style role="ScrollBarTrackSectionHorizontal">
          <states>
            <state name="Normal" backColor="White" backColor2="225, 225, 225" backGradientStyle="Vertical" />
          </states>
        </style>
        <style role="ScrollBarTrackSectionRight">
          <states>
            <state name="Normal" foreColor="White" />
          </states>
        </style>
        <style role="ScrollBarTrackSectionVertical">
          <states>
            <state name="Normal" backColor="White" backColor2="237, 237, 237" backGradientStyle="Horizontal" />
          </states>
        </style>
        <style role="ScrollBarTrackVertical">
          <states>
            <state name="Normal" backColor="White" backColor2="WhiteSmoke" backGradientStyle="Horizontal" />
          </states>
        </style>
        <style role="ScrollBarVertical">
          <states>
            <state name="Normal" backColor="White" backGradientStyle="None" />
          </states>
        </style>
        <style role="SpinButtonDownMaxValue">
          <states>
            <state name="Normal">
              <resources>
                <name>sharedButtonRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="HotTracked" foreColor="White" />
          </states>
        </style>
        <style role="SpinButtonDownNextItem">
          <states>
            <state name="Normal">
              <resources>
                <name>sharedButtonRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="SpinButtonDownNextPage">
          <states>
            <state name="Normal">
              <resources>
                <name>sharedButtonRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="SpinButtonUp">
          <states>
            <state name="Normal">
              <resources>
                <name>sharedButtonRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="HotTracked" foreColor="White" fontBold="True" />
            <state name="Pressed" foreColor="White" />
          </states>
        </style>
        <style role="SpinButtonUpMinValue">
          <states>
            <state name="HotTracked" foreColor="White" />
          </states>
        </style>
        <style role="StateEditorButton">
          <states>
            <state name="Normal" borderColor="DarkGray" />
          </states>
        </style>
        <style role="StatusBarAutoStatusTextPanel">
          <states>
            <state name="Normal" textHAlign="Center" />
          </states>
        </style>
        <style role="StatusBarDatePanel">
          <states>
            <state name="Normal" textHAlign="Center" />
          </states>
        </style>
        <style role="StatusBarMarqueePanel">
          <states>
            <state name="Normal" textHAlign="Center" />
          </states>
        </style>
        <style role="StatusBarPanel">
          <states>
            <state name="Normal">
              <resources>
                <name>sharedButtonRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="StatusBarPanelStateButton">
          <states>
            <state name="Normal" textHAlign="Center" />
          </states>
        </style>
        <style role="StatusBarProgressBar">
          <states>
            <state name="Normal" backColor="237, 237, 237" textHAlign="Center" backGradientStyle="None" />
          </states>
        </style>
        <style role="StatusBarProgressPanel">
          <states>
            <state name="Normal" borderColor="210, 210, 210" />
          </states>
        </style>
        <style role="TabControlClientArea">
          <states>
            <state name="Normal" backColor="Transparent" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="TabControlClientAreaHorizontal">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="TabControlTabItemHorizontalTop" buttonStyle="Flat" />
        <style role="TabControlTabsAreaHorizontalTop">
          <states>
            <state name="Normal" backColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="TabControlTabsAreaVerticalLeft">
          <states>
            <state name="Normal" backColor="237, 237, 237" backGradientStyle="None" />
          </states>
        </style>
        <style role="TabControlTabsAreaVerticalRight">
          <states>
            <state name="Normal" backColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="TabItemHorizontalBottom">
          <states>
            <state name="Normal">
              <resources>
                <name>tabItemHorizontalBottomRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="Selected" foreColor="White" fontBold="True" />
          </states>
        </style>
        <style role="TabItemHorizontalTop" buttonStyle="Flat">
          <states>
            <state name="Normal">
              <resources>
                <name>tabItemHorizontalTopRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="Selected" foreColor="White" fontBold="True">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAigIAAAKJUE5HDQoaCgAAAA1JSERSAAAAVQAAABYIBgAAALE50JgAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAB80lEQVRYR+3YzatpURjH8SWKkvfXUDaFAfI2wEQ7EwyUUEhIQgoDDJj5U+6/4C+73cEdnNHv7mcd1+x0zl33DB/1LT3DT2uvvdYWAARl/DSjH0Y/jWjAfc2AvMhNe1k+QXVj+MvtdqPX6+F6veJ+v3OfGJATeZEb+RnpT0+5Qn8Xi0XGVFxIhEt+5Ph84sXDbrdjv9/jdrtxigan0wkul4tgH7SXvpXLZWy3W1wuF9n5fOYUDNrtNqG+ESparRaWyyWOxyNInFYt9+8G6/X678v9HXU2m2G32+FwOMhVu9lsOAWD56lJoNlsYjQaYbVaSVBatYvFglMweKE2Gg0MBgPM53MJOp1OMZlMOAWDF2q9Xke328V4PJagw+EQ/X6fUzB4odIfp9MJn88nD7Jms5lvU1+7TX3k9H4dZdRvvJabTCZQdHD1+/3weDywWCxyxqkZCAKkx50wQ6GQ3AKsVquE5dQMBF1RbTabBI3FYohEInIroDmnZiC8Xi8cDocETSaT0DQNgUAANOfUDEQ4HJaPfiKRQCaTQTqdRjQaBc05NQMRj8cRDAYlZj6fRy6Xk8A059QMRCqVkvtoNpsFfa2i74I049QNBGHSKq3VatB1HXS7qlarnKIBeYpSqYRCoSBBuf83IE9RqVTQ6XS4bzIgzz8UoyUu5K5g5gAAAABJRU5ErkJgggsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
              <resources>
                <name>tabItemHorizontalTopRubberBlack_Selected</name>
              </resources>
            </state>
            <state name="HotTracked" foreColor="White" />
          </states>
        </style>
        <style role="TabItemVertical">
          <states>
            <state name="Normal">
              <resources>
                <name>tabItemVerticalLeftRubberBlack_Selected</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="TabItemVerticalLeft">
          <states>
            <state name="Normal" foreColor="White" borderColor="Transparent">
              <resources>
                <name>tabItemVerticalLeftRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="Selected" fontBold="True">
              <resources>
                <name>tabItemVerticalLeftRubberBlack_Selected</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="TabItemVerticalRight">
          <states>
            <state name="Normal" foreColor="225, 225, 225">
              <resources>
                <name>tabItemVerticalRightRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="Selected" foreColor="White" fontBold="True" />
          </states>
        </style>
        <style role="TaskPane">
          <states>
            <state name="Normal" backColor="Transparent" borderColor="Transparent" backGradientStyle="None" />
          </states>
        </style>
        <style role="TaskPaneToolbar">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="Active">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="TaskPaneToolbarMenu">
          <states>
            <state name="Normal" foreColor="White" borderColor="Transparent" imageBackgroundStyle="Stretched" backColor2="237, 237, 237">
              <resources>
                <name>scrollBarThumbHorizontalRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ToolbarBase">
          <states>
            <state name="Normal" backColor="White" backColor2="237, 237, 237" backGradientStyle="Vertical" />
          </states>
        </style>
        <style role="ToolbarCloseButton">
          <states>
            <state name="Normal" foreColor="White" borderColor="Transparent" imageBackgroundStretchMargins="3, 12, 3, 4">
              <resources>
                <name>sharedButtonRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="HotTracked" borderColor="Transparent">
              <resources>
                <name>sharedButtonRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ToolbarDockArea">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ToolbarFloatingCaption">
          <states>
            <state name="Normal" backColor="Transparent" borderColor="Transparent" fontName="Tahoma" textHAlign="Center" textVAlign="Middle" imageBackgroundStyle="Stretched" fontSize="8" backGradientStyle="None" imageBackgroundStretchMargins="3, 7, 4, 5">
              <resources>
                <name>scheduleDayOfWeekHeaderRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ToolbarGrabHandle">
          <states>
            <state name="Normal" backColor="237, 237, 237" backGradientStyle="None" />
          </states>
        </style>
        <style role="ToolbarHorizontal">
          <states>
            <state name="Normal" fontName="Tahoma" imageBackgroundStyle="Stretched" fontSize="8" imageBackgroundStretchMargins="1, 3, 1, 3">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ToolbarItem">
          <states>
            <state name="HotTracked">
              <resources>
                <name>sharedButtonRubberBlack_HotTracked</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ToolbarItemButton">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ToolbarItemComboBox">
          <states>
            <state name="Normal" borderColor="Black" />
          </states>
        </style>
        <style role="ToolbarItemFontList">
          <states>
            <state name="Normal" foreColor="30, 30, 30" />
          </states>
        </style>
        <style role="ToolbarItemLabel">
          <states>
            <state name="Normal" foreColor="30, 30, 30" borderColor="Transparent">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ToolbarItemMaskedEdit">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="HotTracked">
              <resources>
                <name>sharedButtonRubberBlack_HotTracked</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ToolbarItemPopupMenu">
          <states>
            <state name="Normal" foreColor="30, 30, 30">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="HotTracked" foreColor="30, 30, 30" borderColor="Transparent" imageBackgroundStyle="Stretched" imageBackgroundStretchMargins="2, 2, 2, 2">
              <resources>
                <name>sharedButtonRubberBlack_HotTracked</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ToolbarItemProgressBar">
          <states>
            <state name="Normal" borderColor="DarkGray" fontName="Tahoma" textHAlign="Center" fontSize="8">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ToolbarItemQuickCustomize">
          <states>
            <state name="Normal" foreColor="White" fontName="Tahoma" textVAlign="Top" imageBackgroundStyle="Stretched" fontSize="8" imageBackgroundStretchMargins="3, 12, 3, 4">
              <resources>
                <name>sharedButtonRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="HotTracked">
              <resources>
                <name>sharedButtonRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="Pressed" backColor="Transparent" borderColor="Transparent" textVAlign="Top" imageBackgroundStyle="Stretched" backGradientStyle="None" imageBackgroundStretchMargins="3, 7, 4, 5" />
            <state name="Active" backColor="Transparent" borderColor="Transparent" backGradientStyle="None" />
          </states>
        </style>
        <style role="ToolbarItemStateButton">
          <states>
            <state name="Normal" foreColor="30, 30, 30" />
            <state name="HotTracked" backColor="237, 237, 237" borderColor="Transparent" backGradientStyle="None">
              <resources>
                <name>sharedButtonRubberBlack_HotTracked</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ToolbarItemTaskPaneLabel">
          <states>
            <state name="Normal" backColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="ToolbarItemTaskPaneMenuDropDownOnly">
          <states>
            <state name="Normal">
              <resources>
                <name>sharedButtonRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="HotTracked">
              <resources>
                <name>sharedButtonRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ToolbarItemTaskPaneNavigation">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ToolTip">
          <states>
            <state name="Normal" backColor="237, 237, 237" foreColor="Black" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="ToolTipBalloon">
          <states>
            <state name="Normal" backColor="237, 237, 237" foreColor="Black" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="ToolTipBalloonTitle">
          <states>
            <state name="Normal" foreColor="Black" />
          </states>
        </style>
        <style role="ToolTipTitle">
          <states>
            <state name="Normal" foreColor="White" />
          </states>
        </style>
        <style role="TreeCell">
          <states>
            <state name="Normal" backColor="Transparent" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="TreeColumnHeader" borderStyle="None">
          <states>
            <state name="Normal" foreColor="White">
              <resources>
                <name>sharedButtonRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="TreeControlArea">
          <states>
            <state name="Normal" backColor="White" backColor2="237, 237, 237" backGradientStyle="Vertical" />
          </states>
        </style>
        <style role="TreeNode">
          <states>
            <state name="Selected" backColor="Gray" foreColor="White" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="TreeNodeExpansionIndicator">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UltraCalculator">
          <states>
            <state name="Normal" borderColor="DarkGray">
              <resources>
                <name>explorerBarGroupItemAreaInnerRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UltraCalculatorButtonAction">
          <states>
            <state name="Normal" foreColor="White" />
          </states>
        </style>
        <style role="UltraCalculatorButtonNumeric">
          <states>
            <state name="Normal" foreColor="White" />
          </states>
        </style>
        <style role="UltraCalculatorDropDown">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UltraCalculatorEditArea">
          <states>
            <state name="Normal" textHAlign="Right" />
          </states>
        </style>
        <style role="UltraColorPicker">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UltraComboEditor">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UltraComboEditPortion">
          <states>
            <state name="Normal" backColor="237, 237, 237" foreColor="30, 30, 30" borderColor="DarkGray" fontName="Tahoma" fontSize="8" backGradientStyle="None" />
          </states>
        </style>
        <style role="UltraCurrencyEditor">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UltraDateTimeEditor">
          <states>
            <state name="Normal" borderColor="DarkGray" />
          </states>
        </style>
        <style role="UltraExpandableGroupBox">
          <states>
            <state name="Normal" backColor="Transparent" backGradientStyle="None" />
          </states>
        </style>
        <style role="UltraFontNameEditor">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UltraGroupBox">
          <states>
            <state name="Normal" backColor="Transparent" backGradientStyle="None" />
          </states>
        </style>
        <style role="UltraGroupBoxContentArea">
          <states>
            <state name="Normal" backColor="White" borderColor="30, 30, 30" backColor2="237, 237, 237" backGradientStyle="Vertical" />
          </states>
        </style>
        <style role="UltraGroupBoxExpansionIndicator">
          <states>
            <state name="Normal" foreColor="White" />
          </states>
        </style>
        <style role="UltraGroupBoxHeader">
          <states>
            <state name="Normal" textVAlign="Top" fontBold="True">
              <resources>
                <name>scheduleDayOfWeekHeaderRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UltraGroupBoxHeaderHorizontalTop">
          <states>
            <state name="Normal">
              <resources>
                <name>headerRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UltraLabel">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UltraMaskedEdit">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UltraNumericEditor">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UltraOptionSet">
          <states>
            <state name="Normal" backColor="237, 237, 237" backGradientStyle="None" />
          </states>
        </style>
        <style role="UltraPictureBox">
          <states>
            <state name="Normal">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UltraProgressBar">
          <states>
            <state name="Normal" backColor="237, 237, 237" foreColor="30, 30, 30" borderColor="DarkGray" textHAlign="Center" backGradientStyle="None" />
          </states>
        </style>
        <style role="UltraStatusBar">
          <states>
            <state name="Normal" backColor="237, 237, 237" backGradientStyle="None">
              <resources>
                <name>baseRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UltraTextEditor">
          <states>
            <state name="Normal" backColor="White" borderColor="DarkGray" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="UltraTimeZoneEditor">
          <states>
            <state name="Normal" borderColor="DarkGray" />
          </states>
        </style>
        <style role="UnpinnedTabAreaBottom">
          <states>
            <state name="Normal">
              <resources>
                <name>explorerBarControlAreaRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UnpinnedTabAreaLeft">
          <states>
            <state name="Normal">
              <resources>
                <name>explorerBarControlAreaRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UnpinnedTabAreaRight">
          <states>
            <state name="Normal">
              <resources>
                <name>explorerBarControlAreaRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UnpinnedTabAreaTop">
          <states>
            <state name="Normal">
              <resources>
                <name>explorerBarControlAreaRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UnpinnedTabItem">
          <states>
            <state name="Normal">
              <resources>
                <name>sharedButtonRubberBlack_Normal</name>
              </resources>
            </state>
            <state name="Selected">
              <resources>
                <name>sharedButtonRubberBlack_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UnpinnedTabItemHorizontalBottom">
          <states>
            <state name="Normal" backColor="Transparent" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="UnpinnedTabItemHorizontalTop">
          <states>
            <state name="Normal" backColor="Transparent" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="UnpinnedTabItemVerticalRight" buttonStyle="FlatBorderless">
          <states>
            <state name="Normal" backColor="Transparent" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="ValueList">
          <states>
            <state name="Normal" backColor="85, 85, 85" borderColor="85, 85, 85" backGradientStyle="None" />
          </states>
        </style>
        <style role="ValueListItem">
          <states>
            <state name="Normal" backColor="237, 237, 237" foreColor="30, 30, 30" borderColor="85, 85, 85" backGradientStyle="None" />
            <state name="Selected" backColor="DarkGray" borderColor="85, 85, 85" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
      </styles>
      <sharedObjects>
        <sharedObject name="ScrollBar">
          <properties>
            <property name="MinimumThumbExtent">20</property>
            <property name="MinimumThumbHeight" colorCategory="{Default}">26</property>
          </properties>
        </sharedObject>
      </sharedObjects>
    </styleSet>
  </styleSets>
  <resources>
    <resource name="baseRubberBlack_Normal" backColor="White" foreColor="30, 30, 30" fontName="Tahoma" fontSize="8" backColor2="225, 225, 225" backGradientStyle="Vertical" />
    <resource name="dropDownButtonRubberBlack_Normal" backColor="Transparent" foreColor="White" borderColor="Transparent" imageBackgroundStyle="Stretched" backGradientStyle="None" imageBackgroundStretchMargins="7, 4, 3, 4">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAFwEAAAKJUE5HDQoaCgAAAA1JSERSAAAACwAAABYIBgAAABbkZmMAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAgElEQVQ4T2P8//8/A9EAqpgVqEEKiE2A2B4Jg/ggcVawOqhiSaDADCD+A8Qgq2AYxJ8GxJLIio2AAj/T0tL+e3h4/NfU1PzPxcUF0/ATKGeErNgOZBoOxSBNdsiKQe7Ep9h+VDFS5Iys0CApIZGUREGJH5TIiUr8xGcrUFolBgMA14xMM12ZutMAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="explorerBarControlAreaRubberBlack_Normal" backColor="237, 237, 237" backColor2="210, 210, 210" backGradientStyle="Vertical" />
    <resource name="explorerBarGroupHeaderRubberBlack_HotTracked" backColor="Transparent" borderColor="Transparent" imageBackgroundStyle="Stretched" backGradientStyle="None" imageBackgroundStretchMargins="3, 12, 3, 4">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAHQEAAAKJUE5HDQoaCgAAAA1JSERSAAAACAAAABQIBgAAALAbfGsAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAhklEQVQoU93RMRIDERiGYRn3wbFWp9TpNKgoOAs3cZgvI5mQ2bU5QMw8qtcw/gdjDORmUUoJGUEp5SLnjNfhsaWULmKMKwgh4Mx7vwJrLXbmFcYY7MxAa42dGSilsDMDKSXOjuNYjxzljhDi/Q//HXDOQcbUvjnn8DEmTGqtuNNaA+m945cn131AmSarZvMAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="explorerBarGroupItemAreaInnerRubberBlack_Normal" backColor="237, 237, 237" backColor2="180, 180, 180" backGradientStyle="Vertical" />
    <resource name="explorerBarItemRubberBlack_Normal" backColor="Transparent" backGradientStyle="None" />
    <resource name="gridCellRubberBlack_HotTracked" backColor="White" foreColor="15, 15, 15" borderColor="Transparent" fontName="Tahoma" imageBackgroundStyle="Stretched" backColorAlpha="UseAlphaLevel" fontBold="True" fontSize="8" backGradientStyle="None" backHatchStyle="None" imageBackgroundStretchMargins="2, 2, 2, 2">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAACgEAAAKJUE5HDQoaCgAAAA1JSERSAAAAKAAAAA4IBgAAAF0LXkYAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAc0lEQVRIS2OcOXNmAgMDQz8QCwDxfyAeLOAD0CFFDEAH/gVih////zMMJgx0UwjIbSAH/hlMDkN2y6gDKY2Z0RAcDUFKQ4BS/aNpcFiHIDB6WUCVCKyqcwdyOCj1MbX0g9wCxBGwNBgPZLwBuXaQYZCb4gD2LAGKQyV8jQAAAABJRU5ErkJgggsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="gridRowRubberBlack_Active" backColor="DimGray" foreColor="White" borderColor="Transparent" fontBold="True" backGradientStyle="None" backHatchStyle="None" />
    <resource name="gridRowRubberBlack_HotTracked" backColor="DarkGray" borderColor="DarkGray" imageBackgroundStyle="Stretched" backGradientStyle="None" backHatchStyle="None" imageBackgroundStretchMargins="2, 2, 2, 2" />
    <resource name="gridRowRubberBlack_Selected" backColor="DarkGray" foreColor="White" borderColor="DarkGray" fontBold="True" backGradientStyle="None" backHatchStyle="None" />
    <resource name="headerRubberBlack_Normal" backColor="Transparent" borderColor="Transparent" imageBackgroundStyle="Stretched" backGradientStyle="None" backHatchStyle="None" imageBackgroundStretchMargins="1, 9, 2, 2">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAADgEAAAKJUE5HDQoaCgAAAA1JSERSAAAACAAAABIIBgAAAGZCn3YAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAd0lEQVQoU92RPQrAIAxGFb2brq6ujm7qpot48K+1P7G0OUEDj0B4ISGRQgjssKG1PuoYY3zovc/GoxmttQ+11iWUUvAm57yEGCM4aEQIARwkeO/BQYJzDhwkWGvxxhizlrzMu0BZKXXe4d+ClBJifu1JSgk388Mb2iHNIp00mo0AAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="mainMenubarHorizontalRubberBlack_Normal" backColor="White" backColor2="225, 225, 225" backGradientStyle="Vertical" />
    <resource name="progressBarFillRubberBlack_Normal" backColor="Transparent" foreColor="White" fontName="Tahoma" textHAlign="Center" imageBackgroundStyle="Stretched" fontSize="8" backGradientStyle="None" imageBackgroundStretchMargins="6, 6, 6, 6">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAA1QAAAAKJUE5HDQoaCgAAAA1JSERSAAAAFAAAABQIBgAAAI2JHQ0AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAPklEQVQ4T2P8//8/A1UByEBqYrhhQFeCnEo2hjlq1MDRMCQhGY0mG3ABQXa2A+kdDUMahCG1ClmqltYgRwEANFB0qcKHZY0AAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="scheduleDayHeaderRubberBlack_Normal" backColor="White" borderColor="DarkGray" backGradientStyle="None" backHatchStyle="None" />
    <resource name="scheduleDayHeaderRubberBlack_Selected" backColor="White" foreColor="30, 30, 30" borderColor="Transparent" fontName="Tahoma" backColor2="165, 165, 165" backGradientStyle="Vertical" />
    <resource name="scheduleDayOfWeekHeaderRubberBlack_Normal" backColor="237, 237, 237" foreColor="White" borderColor="165, 165, 165" fontName="Tahoma" textVAlign="Top" imageBackgroundStyle="Stretched" fontSize="8" backGradientStyle="None" imageBackgroundStretchMargins="4, 12, 4, 4">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAIAEAAAKJUE5HDQoaCgAAAA1JSERSAAAACAAAABQIBgAAALAbfGsAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAiUlEQVQoU93RQQ5DERAGYMLd2NraWtphx0Zch9M40N967aN5T3uASr5IJv/ExFBCCJ62h3N+1FFKuck5j8ajGSmlmxjjCoQQcOW9XwFrLXbmE8YY7MyA1ho7M6CUws4MSClxJYRYQ76TZ2HejLHXP/x3gFIKMrb2yTmH09gwqbXim9YaSO8dvzwALWUBDNGGWfMAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="scheduleDayRubberBlack_Active" backColor="White" backColor2="237, 237, 237" backGradientStyle="ForwardDiagonal" />
    <resource name="scheduleDayRubberBlack_Normal" backColor="White" foreColor="30, 30, 30" borderColor="210, 210, 210" fontName="Times New Roman" fontSize="8" backColor2="237, 237, 237" backGradientStyle="None" backHatchStyle="None" />
    <resource name="scheduleDayRubberBlack_Selected" backColor="White" backColor2="237, 237, 237" backGradientStyle="Horizontal" />
    <resource name="scrollBarArrowDownRubberBlack_Normal" backColor="Transparent" foreColor="White" borderColor="Transparent" imageBackgroundStyle="Stretched" backGradientStyle="None" imageBackgroundStretchMargins="7, 4, 3, 4">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAFwEAAAKJUE5HDQoaCgAAAA1JSERSAAAACwAAABYIBgAAABbkZmMAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAgElEQVQ4T2P8//8/A9EAqpgVqEEKiE2A2B4Jg/ggcVawOqhiSaDADCD+A8Qgq2AYxJ8GxJLIio2AAj/T0tL+e3h4/NfU1PzPxcUF0/ATKGeErNgOZBoOxSBNdsiKQe7Ep9h+VDFS5Iys0CApIZGUREGJH5TIiUr8xGcrUFolBgMA14xMM12ZutMAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="scrollBarArrowUpRubberBlack_Normal" backColor="Transparent" foreColor="White" borderColor="Transparent" imageBackgroundStyle="Stretched" backGradientStyle="None" imageBackgroundStretchMargins="7, 4, 3, 4">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAFwEAAAKJUE5HDQoaCgAAAA1JSERSAAAACwAAABYIBgAAABbkZmMAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAgElEQVQ4T2P8//8/A9EAqpgVqEEKiE2A2B4Jg/ggcVawOqhiSaDADCD+A8Qgq2AYxJ8GxJLIio2AAj/T0tL+e3h4/NfU1PzPxcUF0/ATKGeErNgOZBoOxSBNdsiKQe7Ep9h+VDFS5Iys0CApIZGUREGJH5TIiUr8xGcrUFolBgMA14xMM12ZutMAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="scrollBarThumbHorizontalRubberBlack_Normal" backColor="Transparent" borderColor="Transparent" imageBackgroundStyle="Stretched" backGradientStyle="None" imageBackgroundStretchMargins="7, 4, 3, 4">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAGgEAAAKJUE5HDQoaCgAAAA1JSERSAAAAFgAAAAsIBgAAAHv8fX4AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAg0lEQVQ4T7WSMQ6AIAxFdXDhBnIDVsMMKyfhSB7V0fo/YWiCiyAkbyHpS/vbVUSWKa+KN8h3cIAA4kdYw1o6tuKsYouPE1yAI/TAWjqsFnt83DlnGYEO4LWYo0hKaYg6adBiZirOuSGqODYdG2NkhLeOS8adS9OLbjKedhVz7phh/80DkL0rSuQ04iYAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="scrollBarThumbVerticalRubberBlack_Normal" backColor="Transparent" borderColor="Transparent" imageBackgroundStyle="Stretched" backGradientStyle="None" imageBackgroundStretchMargins="7, 4, 3, 4">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAFwEAAAKJUE5HDQoaCgAAAA1JSERSAAAACwAAABYIBgAAABbkZmMAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAgElEQVQ4T2P8//8/A9EAqpgVqEEKiE2A2B4Jg/ggcVawOqhiSaDADCD+A8Qgq2AYxJ8GxJLIio2AAj/T0tL+e3h4/NfU1PzPxcUF0/ATKGeErNgOZBoOxSBNdsiKQe7Ep9h+VDFS5Iys0CApIZGUREGJH5TIiUr8xGcrUFolBgMA14xMM12ZutMAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="sharedButtonRubberBlack_HotTracked" backColor="WhiteSmoke" borderColor="Transparent" imageBackgroundStyle="Stretched" backGradientStyle="None" backHatchStyle="None" imageBackgroundStretchMargins="2, 2, 2, 2">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAACgEAAAKJUE5HDQoaCgAAAA1JSERSAAAAKAAAAA4IBgAAAF0LXkYAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAc0lEQVRIS2OcOXNmAgMDQz8QCwDxfyAeLOAD0CFFDEAH/gVih////zMMJgx0UwjIbSAH/hlMDkN2y6gDKY2Z0RAcDUFKQ4BS/aNpcFiHIDB6WUCVCKyqcwdyOCj1MbX0g9wCxBGwNBgPZLwBuXaQYZCb4gD2LAGKQyV8jQAAAABJRU5ErkJgggsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="sharedButtonRubberBlack_Normal" backColor="Transparent" foreColor="White" borderColor="Transparent" imageBackgroundStyle="Stretched" fontBold="False" backGradientStyle="None" imageBackgroundStretchMargins="5, 4, 3, 2">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAOwEAAAKJUE5HDQoaCgAAAA1JSERSAAAACQAAABQIBgAAAF/ZF1UAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAApElEQVQoU2NkYGBgBWJRIJYCYm4ghoGvQMYzIH4NEpAE4hlA/AeI/yNhEH8aVJ7BCMj4OXny5P/V1dX/U1JS/kdGRoJxVFTUb6AcSJ7BDqS7rKzsf0RExH8rK6v/Ojo6cAyVZ7AHKQLpNDU1/S8oKIhsJYgNkh9VNDiDABzBeOIOJA9JKjgU/YQlFVCiAyUuvIkOlHxBSdcEmiJASQOEQXyQOCsA1X+uaBzttZ4AAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="spacer" backColor="Transparent" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAqQAAAAKJUE5HDQoaCgAAAA1JSERSAAAAAgAAAAIIBgAAAHK2DSQAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAEklEQVQYV2P8//8/AxiAGCAMAE3WCPnXj6AWAAAAAElFTkSuQmCCCwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="tabItemHorizontalBottomRubberBlack_Normal" foreColor="WhiteSmoke" fontName="Tahoma" imageBackgroundStyle="Stretched" fontSize="8" imageBackgroundStretchMargins="6, 0, 6, 0">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAArAIAAAKJUE5HDQoaCgAAAA1JSERSAAAAVQAAABMIBgAAAOH0QSsAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAACFUlEQVRYR+3Yu64pURgH8LFVCkI0Cl5AUCjI2TqNQlwKtxARl7hlE5ctQUJQi8QDSFQukXgFhVIi6n3eYFdegP+Zb9mcnJPsGOc0irWSf9bMLMXkl29mfUP28vKC2Wwm6PV6wWAwCEqlUhCvCXxIFzidTsLxeBQ+Pz8Fh8MhECDG4zFWqxU2mw12ux0OhwPPAwb7/R7b7Rbr9RrkKchkMjQaDQyHQ0wmE8zncyyXSzbz3DdYLBagTKdTjEYjkCdDDYfDeHt7Q7vdRr/fZ+l0OjwSDLrdLgaDAbMql8u/UcX3ADweDyKRCOLxOKLRKAKBAI8Eg1AodDPz+XwXVHFA3KBgNpths9lgt9thtVphNBp5JBiYTKabm8ViIdALKkXc9aHVaqHRaKBQKK6LfP7yuTr9PcvlcqjVamanUqk46ndQj1znqHeq7hHM6285Kke9vOufPd9WqtPpRDAYRCqVQjabRSaTQSKR4JFgkEwmkcvlkE6nWb9/2/1dLhdisRhbLJVKKBaLDJfnvkE+n2dmhUKB9as3VLfbzaqSvqpqtRqbeaQZ0FcUmdFMVXtFPfv9fvbI0+L7+zvq9TqPRAPyajabzI2edBH1TM3/T6/Xi2q1ilarxfMfBl+V+kGorzqd7lypVNDr9Xj+0YD8yFH0/HH9J7YsHpyevX158vsjUHL8Y7yKZ1S6tPj0/eET3SN5kRv5sfELj1fgOV+yJvQAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==</imageBackground>
    </resource>
    <resource name="tabItemHorizontalTopRubberBlack_Normal" backColor="Transparent" foreColor="225, 225, 225" fontName="Tahoma" imageBackgroundStyle="Stretched" fontBold="False" fontSize="8" backGradientStyle="None" imageBackgroundStretchMargins="6, 10, 6, 10">
      <imageBackground>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</imageBackground>
    </resource>
    <resource name="tabItemHorizontalTopRubberBlack_Selected" backColor="Transparent" borderColor="Transparent" imageBackgroundStyle="Stretched" backGradientStyle="None" imageBackgroundStretchMargins="10, 10, 10, 10">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAALQMAAAKJUE5HDQoaCgAAAA1JSERSAAAAVQAAABYIBgAAALE50JgAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAClklEQVRYR+3YvUuqYRgG8BeaSpLEwSS0tgiLIl1cxEzoG7SSLPEjP5Lwg7Aic5CWoKVRaVK3Al2iJrfaCv+XwxnO0HQdryc/lnPwLRof4eKFR6cf9+N737cCQGHan6l2au38aocHMuoM6EW3qZ5lB9TZPvw9NjYGj8eDfD6Pq6srmQEGdKIX3ejXjrPjKSr0z8LCgsT8ZiERl3507Nx45VGj0SCdTuPy8lLmmwa5XA4Gg4Gwj/wv/VhcXEQymcTZ2ZnI6empjEoDel1cXICobrebqB9ExcrKCg4PD5HNZsWXrFoZdQYnJyc9s4ODg+7L/RM1GAzi+PgYmUxGVO3R0ZGMCoNUKtUzCwQCfVSXy4W9vT3EYjEByqoNh8MyKgzi8bgovmg0Cp/P10d1OBzY2dlBKBQSoBTf39+XUWHA4qMZb7rX6+2j2u12bG1twe/3C9Dd3V1sb2/LqDDomvGmb2xs9FH5stJqtdDr9aKRHRoaktOUumkKw8PD0Ol0wm50dFSi/sQoLlFVVt9XsP+Jyms/MjICo9GIyclJmEwmUc48lxlswCnKbDYLu/HxcWGmTE9PY2JiArOzs7DZbLBarZiZmQHPZQYbzM/P99wsFoswU5aWljA3Nwen04n19XXxBlteXgbPZQYbrK6uYnNzE2tra2BrSjOFLRRB2UZFIhExALDn4rnMYAP2qDRjj882VExVXAiwaeUkdX5+jkKhIJ7d5Yp8fi6Z/heu/bjd4xIqkUiI3ym3t7dikXJzc4O7uzvUajVUKhUZlQb0qlarKJfLuL6+Bj2V+/t7FItFAdpoNPD8/CzzRYOnpyc8PDwIWHoqLy8vKJVKaDabeHt7w/v7u8w3DF5fX1Gv10FPpdVqQeZnDf4C934POC/oA00AAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==</imageBackground>
    </resource>
    <resource name="tabItemVerticalLeftRubberBlack_Normal" imageBackgroundStyle="Stretched" imageBackgroundStretchMargins="10, 6, 10, 6">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAZwIAAAKJUE5HDQoaCgAAAA1JSERSAAAAEwAAAFUIBgAAAKtR370AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAB0ElEQVRYR+1Zy44BURA1O2t+ZP7CSmysCDMd4hGhCSI6COItjIVHgnhEwk4vrHzFfNKZrjIxJnGNzeyqk5NOuqvPfeT0vXVPvQCw0RWNRt+tm27hlR/8cdntdpvT6WQ4HA6+24gsEomYHo8H4XAYFimCwSC8Xi9cLpcSbrcbPp+P43O5HGq1GhN9WMB4PEa73UalUkE2m0U8Hoff71dC0zTous7x/X4f8/mcydDpdHA8HvnBYDBAvV5HoVDgVlVIpVJMRPHL5ZK/Z7LNZgPTNLHdbjGdTtHtdjkwmUwqQUNrNpscv9/vcT6fL2Sn0+lKRsMlsnK5jEQioQRNxVNkrVYLxWIR1JAKNEyadGr8Yc+E7DqHMme/RSvSEGko1jORhkhDpCGbMP8Fsm/KvqlIRUUaIg2RhhzELsdqOoxK5iiZo2SOkjlK5njHJpTFURbH/1gcb73tyWTC7rthGIjFYkqk02k21Cn+cDj8eNvkupNRvtvtMJvN0Ov1njLKSdwUf0vG9YDVaoXFYoHRaIRGo8HeNtUEVKCeVatVjl+v19yZa6WCigWlUokD8vk8u+2BQECJUCiETCbD8cPhkGsCTPZd+nizevj5yGm/fXfvuPMFK12oah+MHBEAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==</imageBackground>
    </resource>
    <resource name="tabItemVerticalLeftRubberBlack_Selected" imageBackgroundStyle="Stretched" imageBackgroundStretchMargins="10, 6, 10, 6">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAgwIAAAKJUE5HDQoaCgAAAA1JSERSAAAAEwAAAFUIBgAAAKtR370AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAB7ElEQVRYR+2ZyWoCQRCGFUXw4I7LwYO4o7hcFEFxBRdQQUUEF1xxOeWkIl58gzxKHiWP9MfqiTESOznmUgM/goN/tzNV3dVfqQGo6FKr1bPrx8tVSfHFH5dOp1O53e4HqT7N3q43kUgkkEqlEAwG4XA4aBSptFotXC4XotEoCoUCut2umNUr/WgymWAwGKBWqyGdTiMQCMBgMEhlsVgQCoWQy+XQ6XSw2WyEGTKZDE6nE1arlRihWCwiFovBZrNJRbOif1KtVsVE9vu9YjabzYTZcrlEv99HuVxGPB4HjS6T0+lEMplEo9HAfD7H8XhUzC6XizBbLBZotVrIZrMIh8PQ6/VSmUwmRCIRVCoVjEYjHA6Hn2bNZlM8M4/H8+sLoBfm9XqRz+cxHA7ZjJ/ZLV85NO65yen0sIpwaHBoSHZ1Dg0ODQ4NLqmUmpb3Td43+VDBB7GvLODygMsDLg+4PODy4EkW8OLIi+O/Lo7f2Xa73Rbg3OfzQaPRSEWo2u/3C6g+Ho/vBPlG3bfbrQDlpVJJUHez2SyV3W4X1L1er4vzw/l8vvcDCN+v12v0ej1hRtT9L4RP1J36B9PpVKHut06F1WoV6J6QPM2UGgdGo1EqGojIPIFymsBut1PMPg2nV2z6/lur4/u9Z4n+ATUWYLVifyWMAAAAAElFTkSuQmCCCwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA</imageBackground>
    </resource>
    <resource name="tabItemVerticalRightRubberBlack_Normal" imageBackgroundStyle="Stretched" imageBackgroundStretchMargins="10, 6, 10, 6">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAApAIAAAKJUE5HDQoaCgAAAA1JSERSAAAAEwAAAFUIBgAAAKtR370AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAACDUlEQVRYR+1ZuaoiQRTtQVFE3BBxCURFUDRQUBP3xA1E3DBwx0ARTF1Q8QuExiUV3h/MN/gFMx/hV5ic6SrnOfim3/DCCW7DCbq7+tTtrttVp879djwecbvdhI+43+/CF46fUpsDgDfedjweo16vI5PJIBgMwmazQalUQrolC61WC6/Xi2QyiUQiAbVazdp9lwgFYbfbYTqdotFoIJ1Ow+/3w2w2Q6/Xy8LhcCAWi6FarWI2m2G5XL53KgqiKPILg8EAhUIB4XCYR8cI5eB2u3mn3W4X6/Ual8sFzWbzQXg6nbDZbDAajVAsFhEKhWCxWGA0GmXhdDo5Wb/f588xssPh8CBjA7BYLNDpdJDNZvn30Gg0UCgUsmDRstdstVrPyK7X6ytZu93mPXo8HqhUqk8HgEUciUSITBoA+mYvaUKp8effpNSg1HisATTT0oIirdo0OdLk+ImmpdSg1KDUoE3FY1dHypGUIylHsiKefwHJA5IHJA9IHvyv8oB528zwHg6HyOVyCAQCMBgM3JKWg91uRzwe5/b1X972+Xx+GuWlUom77larFSaTSRYul4vXDphRzuz/F9d9v99jPp+j1+shn89z1/1fFj4jS6VSXIKtVitO+LsQIQrb7RaTyQS1Wo1XH3w+H49Ip9PJgr1mNBpFpVJBuVzmNYNnpeLdwv+q2PtQDvkhnQ9YyYPhFyTfbBuNW2o6AAAAAElFTkSuQmCCCwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA</imageBackground>
    </resource>
  </resources>
</styleLibrary>