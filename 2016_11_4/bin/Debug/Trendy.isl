<?xml version="1.0" encoding="utf-8"?>
<styleLibrary>
  <annotation>
    <lastModified>2007-09-24T15:08:36</lastModified>
  </annotation>
  <styleSets defaultStyleSet="Default">
    <styleSet name="Default" buttonStyle="FlatBorderless" viewStyle="Standard" useOsThemes="False" useFlatMode="True">
      <componentStyles>
        <componentStyle name="Inbox Form">
          <properties>
            <property name="BackColor" colorCategory="{Default}">White</property>
          </properties>
        </componentStyle>
        <componentStyle name="Inbox MonthCalendar">
          <properties>
            <property name="TitleBackColor" colorCategory="{Default}">112, 188, 18</property>
          </properties>
        </componentStyle>
        <componentStyle name="Inbox Panel">
          <properties>
            <property name="BackColor" colorCategory="{Default}">White</property>
          </properties>
        </componentStyle>
        <componentStyle name="UltraDockManager" buttonStyle="FlatBorderless" useFlatMode="True">
          <properties>
            <property name="GroupPaneTabStyle" colorCategory="{Default}">Flat</property>
            <property name="UnpinnedTabStyle" colorCategory="{Default}">Flat</property>
          </properties>
        </componentStyle>
        <componentStyle name="UltraGrid">
          <properties>
            <property name="AddNewBoxButtonConnectorStyle" colorCategory="{Default}">None</property>
          </properties>
        </componentStyle>
        <componentStyle name="UltraPictureBox" resolutionOrder="None" />
        <componentStyle name="UltraTabbedMdiManager">
          <properties>
            <property name="TabStyle" colorCategory="{Default}">Flat</property>
          </properties>
        </componentStyle>
        <componentStyle name="UltraTabControl">
          <properties>
            <property name="Style" colorCategory="{Default}">Flat</property>
          </properties>
        </componentStyle>
        <componentStyle name="UltraTabStripControl">
          <properties>
            <property name="Style" colorCategory="{Default}">Flat</property>
          </properties>
        </componentStyle>
        <componentStyle name="UltraTree" buttonStyle="FlatBorderless" headerStyle="Standard" useFlatMode="True" />
      </componentStyles>
      <styles>
        <style role="Base">
          <states>
            <state name="Normal" foreColor="DimGray" fontName="Segoe UI" textVAlign="Middle" fontSize="8" themedElementAlpha="Transparent" />
            <state name="HotTracked" fontName="Segoe UI" />
          </states>
        </style>
        <style role="Button">
          <states>
            <state name="Normal">
              <resources>
                <name>buttonTrendy_Normal</name>
                <name>headerTrendy_normal</name>
              </resources>
            </state>
            <state name="Selected" imageBackgroundStretchMargins="3, 3, 3, 3" />
          </states>
        </style>
        <style role="CalendarComboDateButton">
          <states>
            <state name="Normal" foreColor="White">
              <resources>
                <name>menuItembuttonTrendy_hotTracked</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="CalendarComboDateButtonArea">
          <states>
            <state name="Normal" backColor="White" borderColor="Gainsboro" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="CalendarComboDropDown">
          <states>
            <state name="Normal" backColor="White" borderColor="Gainsboro" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="ComboDropDownButton">
          <states>
            <state name="Normal" foreColor="White" imageBackgroundStyle="Stretched" imageBackgroundStretchMargins="1, 1, 1, 1" />
            <state name="HotTracked" backColor="117, 200, 255" foreColor="White" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
            <state name="Pressed" backColor="117, 200, 255" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="ControlArea">
          <states>
            <state name="Normal" borderColor="Gainsboro" />
          </states>
        </style>
        <style role="DayViewAllDayEventArea">
          <states>
            <state name="Normal" backColor="249, 252, 252" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="DayViewTimeSlotDescriptor">
          <states>
            <state name="Normal">
              <resources>
                <name>dayViewTimeSlotDescriptorTrendy_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="DayViewTimeSlotNonWorkingHour">
          <states>
            <state name="Normal" backColor="216, 248, 181" borderColor="LightGray" backGradientStyle="None" backHatchStyle="None" />
            <state name="Selected" borderColor="Gainsboro">
              <resources>
                <name>gridCellTrendy_hotTracked</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="DayViewTimeSlotWorkingHour">
          <states>
            <state name="Normal" backColor="WhiteSmoke" borderColor="Gainsboro" backGradientStyle="None" backHatchStyle="None" />
            <state name="Selected" borderColor="Gainsboro">
              <resources>
                <name>gridCellTrendy_hotTracked</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="DesktopAlertButton">
          <states>
            <state name="Normal" backColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
            <state name="HotTracked" backColor="Transparent" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="DesktopAlertCloseButton">
          <states>
            <state name="Normal" backColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
            <state name="HotTracked" borderColor2="Transparent">
              <resources>
                <name>headerTrendy_normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="DesktopAlertControlArea">
          <states>
            <state name="Normal">
              <resources>
                <name>progressBarFillTrendy_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="DesktopAlertDropDownButton">
          <states>
            <state name="Normal">
              <resources>
                <name>spacer</name>
              </resources>
            </state>
            <state name="HotTracked">
              <resources>
                <name>headerTrendy_normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="DesktopAlertGripArea">
          <states>
            <state name="Normal">
              <resources>
                <name>scheduleDayHeaderTrendy_Selected</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="DesktopAlertLink">
          <states>
            <state name="Normal" foreColor="White" textVAlign="Top" />
            <state name="HotTracked" foreColor="White" textVAlign="Top" />
          </states>
        </style>
        <style role="DesktopAlertPinButton">
          <states>
            <state name="Normal" foreColor="White">
              <resources>
                <name>spacer</name>
              </resources>
            </state>
            <state name="HotTracked">
              <resources>
                <name>headerTrendy_normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="DockAreaSplitterBarHorizontal">
          <states>
            <state name="Normal" backColor="249, 252, 252" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="DockAreaSplitterBarVertical">
          <states>
            <state name="Normal" backColor="White" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="DockControlPaneContentArea">
          <states>
            <state name="Normal" backColor="249, 252, 252" borderColor="White" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="DockFloatingWindowCaptionHorizontal">
          <states>
            <state name="Normal" foreColor="White" />
            <state name="Active" backColor="153, 204, 51" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="DockPaneCaptionButton">
          <states>
            <state name="Normal" borderColor="Transparent">
              <resources>
                <name>shareTransparentImageTrendy</name>
              </resources>
            </state>
            <state name="HotTracked" foreColor="Black" />
          </states>
        </style>
        <style role="DockPaneCaptionHorizontal">
          <states>
            <state name="Normal" backColor="249, 252, 252" backGradientStyle="None" backHatchStyle="None" />
            <state name="Active" backColor="153, 204, 51" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="DockSlidingGroupHeaderHorizontal">
          <states>
            <state name="Normal" foreColor="White">
              <resources>
                <name>headerTrendy_normal</name>
              </resources>
            </state>
            <state name="HotTracked" borderColor="Transparent" />
          </states>
        </style>
        <style role="DropDownButton">
          <states>
            <state name="Normal" borderColor="White">
              <resources>
                <name>dropDownButtonTrendy_Normal</name>
              </resources>
            </state>
            <state name="HotTracked" borderColor="White" />
            <state name="Pressed" backColorAlpha="Opaque" />
          </states>
        </style>
        <style role="DropDownControlArea">
          <states>
            <state name="Normal" backColor="White" borderColor="Gainsboro" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="DropDownEditorButton" buttonStyle="FlatBorderless">
          <states>
            <state name="Normal">
              <resources>
                <name>editorButtonTrendy_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="EditorButton">
          <states>
            <state name="Normal">
              <resources>
                <name>editorButtonTrendy_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ExpansionIndicator">
          <states>
            <state name="Normal" backColor="Transparent" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="ExplorerBarControlArea">
          <states>
            <state name="Normal">
              <resources>
                <name>controlAreaTrendy_normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ExplorerBarGroupHeader">
          <states>
            <state name="Normal" foreColor="White" borderColor="Transparent" fontBold="True">
              <resources>
                <name>headerTrendy_normal</name>
              </resources>
            </state>
            <state name="HotTracked" borderColor="Transparent" />
          </states>
        </style>
        <style role="ExplorerBarGroupItemAreaInner">
          <states>
            <state name="Normal">
              <resources>
                <name>ExplorerBarGroupItemAreaInnerTrendy_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ExplorerBarItem">
          <states>
            <state name="Normal" foreColor="66, 113, 11" />
          </states>
        </style>
        <style role="ExplorerBarNavigationOverflowButtonArea">
          <states>
            <state name="Normal" foreColor="White">
              <resources>
                <name>menuItembuttonTrendy_hotTracked</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ExplorerBarNavigationOverflowQuickCustomizeButton">
          <states>
            <state name="Normal" foreColor="White" />
          </states>
        </style>
        <style role="ExplorerBarNavigationSplitter">
          <states>
            <state name="Normal" backColor="0, 153, 255" foreColor="White" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="GridAddNewBox">
          <states>
            <state name="Normal">
              <resources>
                <name>gridAddNewBoxTrendy_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="GridAddNewBoxButton" buttonStyle="Button">
          <states>
            <state name="Normal" foreColor="White" fontBold="True" imageBackgroundStretchMargins="4, 4, 4, 4" />
          </states>
        </style>
        <style role="GridAddNewBoxPrompt">
          <states>
            <state name="Normal" foreColor="White" fontBold="True" fontSize="13" />
          </states>
        </style>
        <style role="GridBandHeader">
          <states>
            <state name="Normal">
              <resources>
                <name>headerTrendy_normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="GridCaption">
          <states>
            <state name="Normal" backColor="White" fontBold="True" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="GridCardArea">
          <states>
            <state name="Normal" backColor="White" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="GridCardCaption">
          <states>
            <state name="Normal">
              <resources>
                <name>gridCardCaptionTrendy_Normal</name>
              </resources>
            </state>
            <state name="Selected" backColor="153, 204, 51" foreColor="Black" borderColor="White" backGradientStyle="None" backHatchStyle="None" />
            <state name="Active" foreColor="Black" />
          </states>
        </style>
        <style role="GridCell">
          <states>
            <state name="Normal" foreColor="45, 45, 45" borderColor="White" fontSize="8" />
            <state name="Selected" borderColor="Transparent" fontBold="True" />
            <state name="HotTracked">
              <resources>
                <name>gridCellTrendy_hotTracked</name>
              </resources>
            </state>
            <state name="Active" backColor="Transparent" foreColor="Black" borderColor="Transparent" fontSize="8" backGradientStyle="None" backHatchStyle="None" />
            <state name="EditMode">
              <resources>
                <name>gridCellTrendy_EditMode</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="GridColumnHeader">
          <states>
            <state name="Normal" foreColor="White" textHAlign="Left" fontBold="True" fontSize="9" imageBackgroundStretchMargins="6, 6, 6, 6">
              <resources>
                <name>headerTrendy_normal</name>
              </resources>
            </state>
            <state name="HotTracked">
              <cursor>Hand</cursor>
            </state>
          </states>
        </style>
        <style role="GridControlArea">
          <states>
            <state name="Normal" backColor="White" borderColor="White" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="GridExpansionIndicator">
          <states>
            <state name="Normal">
              <resources>
                <name>gridExpansionIndicatorTrendy_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="GridGroupByBox">
          <states>
            <state name="Normal" backColor="White" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="GridGroupByBoxPrompt">
          <states>
            <state name="Normal" backColor="Transparent" foreColor="SteelBlue" fontSize="12" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="GridHeader" borderStyle="None" />
        <style role="GridRow">
          <states>
            <state name="Normal" borderColor="LightGray" />
            <state name="Selected">
              <resources>
                <name>gridRowTrendy_Selected</name>
              </resources>
            </state>
            <state name="Active">
              <resources>
                <name>gridRowTrendy_Active</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="GridRowScrollRegionSplitBox">
          <states>
            <state name="Normal" backColor="153, 204, 51" borderColor="White" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="GridRowSelector">
          <states>
            <state name="Normal">
              <resources>
                <name>gridRowSelectorTrendy_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="GridRowSelectorHeader">
          <states>
            <state name="Normal" backColor="White" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="GroupPaneTabItemAreaVerticalRight">
          <states>
            <state name="Normal" backColor="249, 252, 252" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="GroupPaneTabItemVerticalRight">
          <states>
            <state name="Normal" textVAlign="Bottom" />
          </states>
        </style>
        <style role="Link">
          <states>
            <state name="HotTracked" textVAlign="Top" />
          </states>
        </style>
        <style role="ListViewColumnHeader">
          <states>
            <state name="Normal" backColor="Transparent" foreColor="White" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="ListViewControlArea">
          <states>
            <state name="Normal" backColor="249, 252, 252" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="ListViewGroupHeader">
          <states>
            <state name="Normal" foreColor="White" borderColor="Gainsboro">
              <resources>
                <name>headerTrendy_normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ListViewItem">
          <states>
            <state name="Selected" backColor="0, 153, 255" borderColor="White" imageBackgroundStyle="Stretched" backGradientStyle="None" backHatchStyle="None" imageBackgroundStretchMargins="3, 3, 3, 3" />
            <state name="HotTracked">
              <resources>
                <name>listViewItemTrendy_HotTracked</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="MainMenubar">
          <states>
            <state name="Normal">
              <resources>
                <name>mainMenubarTrendy_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="MenuItem">
          <states>
            <state name="Normal" backColor="250, 252, 253" backGradientStyle="None" backHatchStyle="None" />
            <state name="HotTracked" foreColor="White">
              <resources>
                <name>toobarItemPopupMenuTrendy_HotTracked</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="MenuSideStrip">
          <states>
            <state name="Normal" foreColor="White" backGradientStyle="Horizontal">
              <resources>
                <name>headerTrendy_normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="MenuTearAwayStrip">
          <states>
            <state name="Normal" backColor="127, 194, 47" borderColor="White" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="MonthViewMultiScrollButton">
          <states>
            <state name="Normal" backColor="Transparent" foreColor="White" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="NavigationBarLocationButton">
          <states>
            <state name="Normal" backColor="Transparent" foreColor="30, 30, 30" borderColor="Transparent" fontBold="False" backGradientStyle="None" backHatchStyle="None">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAA+QAAAAKJUE5HDQoaCgAAAA1JSERSAAAAKAAAACgIBgAAAIz+uG0AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsRAAALEQF/ZF+RAAAAYklEQVRYR+3SsQ3AMAwDwXj/oRU4K3wKFWfAJQHixDMzz+p3C27+q8t9192sp+Af13HiqkiQYBWoeRskWAVq3gYJVoGat0GCVaDmbZBgFah5GyRYBWreBglWgZq3QYJVoOZfWTs5xeZ58m0AAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
            </state>
            <state name="HotTracked">
              <resources>
                <name>scheduleOwnerHeaderTrendy_Normal</name>
              </resources>
            </state>
            <state name="Pressed">
              <resources>
                <name>scheduleOwnerHeaderTrendy_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="NavigationBarLocationListItem">
          <states>
            <state name="HotTracked">
              <resources>
                <name>scheduleOwnerHeaderTrendy_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="NavigationBarLocationTextButton">
          <states>
            <state name="Normal" backColor="Transparent" backGradientStyle="None" backHatchStyle="None">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAA+QAAAAKJUE5HDQoaCgAAAA1JSERSAAAAKAAAACgIBgAAAIz+uG0AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsRAAALEQF/ZF+RAAAAYklEQVRYR+3SsQ3AMAwDwXj/oRU4K3wKFWfAJQHixDMzz+p3C27+q8t9192sp+Af13HiqkiQYBWoeRskWAVq3gYJVoGat0GCVaDmbZBgFah5GyRYBWreBglWgZq3QYJVoOZfWTs5xeZ58m0AAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
            </state>
          </states>
        </style>
        <style role="NavigationBarPreviousLocationsDropDownButton">
          <states>
            <state name="Normal" foreColor="135, 135, 135">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAA+QAAAAKJUE5HDQoaCgAAAA1JSERSAAAAKAAAACgIBgAAAIz+uG0AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsRAAALEQF/ZF+RAAAAYklEQVRYR+3SsQ3AMAwDwXj/oRU4K3wKFWfAJQHixDMzz+p3C27+q8t9192sp+Af13HiqkiQYBWoeRskWAVq3gYJVoGat0GCVaDmbZBgFah5GyRYBWreBglWgZq3QYJVoOZfWTs5xeZ58m0AAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
            </state>
            <state name="HotTracked" backColor="Transparent" foreColor="44, 165, 67" backGradientStyle="None" backHatchStyle="None">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAA+QAAAAKJUE5HDQoaCgAAAA1JSERSAAAAKAAAACgIBgAAAIz+uG0AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsRAAALEQF/ZF+RAAAAYklEQVRYR+3SsQ3AMAwDwXj/oRU4K3wKFWfAJQHixDMzz+p3C27+q8t9192sp+Af13HiqkiQYBWoeRskWAVq3gYJVoGat0GCVaDmbZBgFah5GyRYBWreBglWgZq3QYJVoOZfWTs5xeZ58m0AAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
            </state>
            <state name="Pressed" backColor="Transparent" foreColor="0, 109, 164" backGradientStyle="None" backHatchStyle="None">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAA+QAAAAKJUE5HDQoaCgAAAA1JSERSAAAAKAAAACgIBgAAAIz+uG0AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsRAAALEQF/ZF+RAAAAYklEQVRYR+3SsQ3AMAwDwXj/oRU4K3wKFWfAJQHixDMzz+p3C27+q8t9192sp+Af13HiqkiQYBWoeRskWAVq3gYJVoGat0GCVaDmbZBgFah5GyRYBWreBglWgZq3QYJVoOZfWTs5xeZ58m0AAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
            </state>
          </states>
        </style>
        <style role="NavigationToolbarBackButton">
          <states>
            <state name="Normal">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAXwIAAAKJUE5HDQoaCgAAAA1JSERSAAAAGQAAABkIBgAAAMTphWMAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAA7BAAAOwQG4kWvtAAAByElEQVRIS92Vv0sCYRzG6w+IHNuktUWhH1RUWmgIQQkN0aRLQ2UZNZWBhUMN0RA0FJQOQVBQRFJBkCRKIqG7BQ4NjRf9WGr49ryi13n3vtdpuCR8QN97nvdz33sPrCeiupp/mKTW1FxQeFJGpthKt1mBHawoYL+tRvpCCTYwgQiQAOmQL4pNIiFXspFqdwMJUAXkkeVOppGEkh1eQFUioacRlUkCt512QL8RSjr0MhL6ZY+uTDJ/0xUDJCIQd1L6OUrbmSlhptgNK89Hlkxed1sBiVhOjNLTaw5dos37aWGu2Je4Eu9Vjx8Qj63MIn18vRUE7LOW9nFzqq6tJJInGYv2BgGpOc7ty5uXvqzezWhynO6IRjJ81hcEpGYpMUvvnz9TMBFb42VVa9pJnCd2DyAe7vMhenx5kCdaiPu5OVXXopmk56jfDEiPi/xlQeSLzenmsAf/4Jm19XAgDEiPYGqdxi8ndDPoB7lvF1tsOXCYgQToD2TRbRRK2IXmiNMCJEBVwHryWWjORGlu2hu0gCygCoghWzaBrqR00bTr8oBTQAKk4nX5dVXerCGJstCw4zIDmwLuXf9JwisbXTP092t0M1Hu/0i+Aah8LXEF9F/LAAAAAElFTkSuQmCCCwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA</imageBackground>
            </state>
            <state name="HotTracked">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAANgIAAAKJUE5HDQoaCgAAAA1JSERSAAAAGQAAABkIBgAAAMTphWMAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAA7BAAAOwQG4kWvtAAABn0lEQVRIS92VvUtCURjGrT+gnFscGxXas7mllogmg/YcGxpuTUEEDU6KoIuTQ0XgFFhTDYUhRB8kEkU0BCciaih4ek7q7d57ProqLh34DZ7zPud333MOOAQgMvAhJYNm4IKfkwrTxf4zEmSKrHmQvxNh8kYJN4iSIhEEFpptcdQk1ErKT5glgqALmqzVdqZICvdYJOgRwZwi8kkyDUwR/EWuaa0RzPuOzifZuESVwMT2DVB/AUp35pp2tuC9H1eyWkeCwMTWFfD4wShHnt3aarkmtJL0KdIEOvK3wPtXSyBH5lpfF8gmOyK3k6VjOARB9vgIgmPzQq3TZGcUyfwRHIIg6+fA26dfI+d0tYE5tZPpA6QIdMwdAo3XX9HKmb4ukI0rnUxWECOwUXloiZZP7HXcQ3/x0jqxiwKBDYddLFTtNcw72tclJ8fLiBFB0Ac1ZkeNErkwVkKcCIIekDn3LpQ78ZpHioiTGkEXVFnr68Aq6SwO55AiOwQGRHvdfa7ejw0l8Z1rFrFIFkkP2q/uS6ILh50L9fcbdjNT3f+RfANccB6WMXyERgAAAABJRU5ErkJgggsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
            </state>
            <state name="Pressed">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAWAIAAAKJUE5HDQoaCgAAAA1JSERSAAAAGQAAABkIBgAAAMTphWMAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAA7BAAAOwQG4kWvtAAABwUlEQVRIS92VvUtCURyG6w+IXJvcWpX+gHJuqaVo0xobagkaAnWTIChqs+I6VEtIhOHgok1ZJEofQoFIRNqHdYZEsA9+vUf0cj/OuV0Tly48oOe+73nuOefC7SWinq5fXNJtui5o7JSdVcQq5AYeENTA/7vt9KUSTOAAEcAAWVBsih0yoVCyX/ocBwxQGxSRFa7MJFEKVR+gP8LQM4l0ko3LNw+g3wjnmVWGoa/bOp0kdFpOApKxev5EFy812s2/SjPNrqI9H1WylLpzA5Kxkn6g0nsdXaKt3KM01+wzoWQufjsPSMRmpkS1j++GgF/r6XthztAdaYnUlcxErwKAjBzmn9XJWz+Wj4umnKA7ZpJM7mQCgIwEEzdUrX/pRHxMlDWMmVcyGj7xAhIxETmjQqWqihZj18KcoesyrWR4LeUEZEX8utwQze3nLHOYQ3zw3DoUSiiArPAfXdHU9ollBv2A8O3ig4P+mBMwQB2QRbdfKuE3BhaiLsAA/QHeU8/CdCZac9/sngtkAbVBElndCiwlrZu90xEvOAAkgTXvq6+r9mFtSXT76lOcPT5lRIPwqTuSiMp2x2x9fu1OJsv9H8kPUb/OCJo94z8AAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==</imageBackground>
            </state>
          </states>
        </style>
        <style role="NavigationToolbarButtonArea">
          <states>
            <state name="Normal">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAsgIAAAKJUE5HDQoaCgAAAA1JSERSAAAAQQAAAB8IBgAAALhzoMIAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAA7BAAAOwQG4kWvtAAACG0lEQVRoQ+1ZW27CMBCkl0A9RNtb0gPQk/SvP3zhU5RH4yQECoGUAgnkAd1xY6kCkpifKk6CZCGBiXbHM8t69u50OrVq/wIIdV+1B0AoIY8F74NxC8tdru6ns/mzZTuvo7HRp89YWRfiQ5yIF3HLHPLyvAqC/CHenenny2jM+6Y1YZ63ZocwZMfjkdFDS7cQF+JDnKblMMSN+P/mcw2MTBCAomlO3rhps/1+X7qEVQ7BDwJmcIshD8kKJRAkaga3e4SilsmfAzRxZgSG3cuSxgUTpAQsolMcx5UAIYpisIFJaZyz4QKEhbt8QtELAj0lkCUT3w9EMUd+hSBQVe2APiqa020P6hv+NQpBQBHZbLaVBMHzvkSRLAQBlIGGdDtllXgD+pdDfkogqDxQxz1JkjQg4OBqz4QGhLTNb5jQyOH3wld7JuCWqQxCFEWV7BMOh1ANBHRU6+9NJUFQ7hhxdyBnpnIgQAowhpBfYcdIt6zHwfCDocXUsSvMinm384UUXHf1UAhC6id0Od2/y2qj3Xo4qHGc2/ATushPCQRsJFuqB1mEod5FEvHDUEE+ys4SUEod5jaKJB4AQ0I3ViDe7XbHwOjUY2xfY0Gm5X7pNht98uigJ+E44TZ2KyX/Y38cJwzany+W8BTJbTbgNgsJ5FnvqnOHtpw7DEflnjsgvnTu0CGHWZx+FgNkbWgmUEUTqLypTZW+a5hATPgBUzaCg/xJgZgAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==</imageBackground>
            </state>
          </states>
        </style>
        <style role="NavigationToolbarForwardButton">
          <states>
            <state name="Normal" imageBackgroundStyle="Centered">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAOwIAAAKJUE5HDQoaCgAAAA1JSERSAAAAGQAAABkIBgAAAMTphWMAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAA7BAAAOwQG4kWvtAAABpElEQVRIS83Vv0tCURQH8PoDIsc2aW1R6AcVlq/IMohqa1RqiKag1cEgamoIN4PQIQoakiCiiLKCQBp0aAhaHAqKljtYDTWcvld88uy+e54/EnrwAX33nPN9970nthJRS9MPGdJsTQ8o3qlqdxG76/GCAasW8rvXaQYbggEuSIIAYuRLwS67QG3IZqZ3FgRQDfKoVXZmG7J22xcGqpNAX0WQEhK57jeAGiTQX751SsjK5UAaSOe58EiRm4B23dKXMJ9PRcjS+aAXiING+vwu0HomxNZhhrANCZ/6loE4MsQ89h5ibC3m+JXfydzxUBSIYw2Rn6+eTmj+bFLXM6OETB8NR4E41pD3rwJtZTe4enUngUMjBMQxQ14/XmjxYoGtxRyPshPfwYgbiCNDsm85mkhNsXWYYf/gZWr3/mgCSCd+n9Su/eqJ2r5d8mTX7pgbBFADcuht14bIhc5kwAMCqA6yr/gs2BC52LEz7oEcUA3SqC3vwDHELHBtB0OQAtIQpfXi62qn6j8t2dwWD7rBb6FcdcMhuit1Ol/TTpyG/cnt+tchP+k8LGg6KqeWAAAAAElFTkSuQmCCCwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA</imageBackground>
            </state>
            <state name="HotTracked">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAGQIAAAKJUE5HDQoaCgAAAA1JSERSAAAAGQAAABkIBgAAAMTphWMAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAA7BAAAOwQG4kWvtAAABgklEQVRIS82VP0vDQBiH1Q+gnV0yOqbgbrvqoptjBUeRfgCHuDk6KEJd0sWpoILgIEIRBAWFgIsiSEGXIsotQgeFn78LTUm53JumMWDggSTvnyfv3ZVOApgo/NKSoilcEK7UqFOcfaBMqmQ7hn4up/UQJWxQIk2iCAQ6fXEpSWiVtLpYIYogAx3mGpMlSvw3rBGMiWLdkMiQ7L2gSpATxfrB0hmSnUe0CWx0e8Dusz0eq/Oj/RmSbD2gTCDBQvR+gH2KUnJVoqR+jzqBhJZE1/GrnMs+FeN3sn4Dj0AiLtH31+/A5p21ZtmQrF7BI5CIS76+gYMnMd+cZOkSNQKJSKIPwMatnMs+rjHJwjkcAgktCT6BxQs5jz2SN15b50/hE9ho8HhL8VjMSzxd+uVcCw5RBDkIWDtjlejA7BFcogjGQNeFeyFKdHC6CZcEBBloM3cwQaokSpg6RI2cEFhQ/Xh4XJMY+U8rPIoNOKQSw/jq3BLbl6a9zzRJWrM/Wa5/LfkF40YeDtiNi+AAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==</imageBackground>
            </state>
            <state name="Pressed">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAANQIAAAKJUE5HDQoaCgAAAA1JSERSAAAAGQAAABkIBgAAAMTphWMAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAA7BAAAOwQG4kWvtAAABnklEQVRIS82VTUsCURSG8wdEblu5a6v0A9JtbWrVVmtZUZt2QbZzWVBRtBkX0SYsSGjRRqEoAkOiKIhCI0iL6EIYfVic3js4MnbnnvEjoYEHZuZ8PHPuvaKHiDrafklJu2m7wFypeqdIPlEAhMCcDfkccOvBStDAC+JAAGLIVcReJ6FWsnlfHgICUAPkkKtM5igxbkoRQE0iUFcjUiSLZ88hQC0iUF9dOkUSOy6kAOkovn7SfOZBG7fVGdb+1Ehm0rcBQBwopLfyNy1lCmweeghHyeTu1RQgDimxrq2LRzYXfYLK72Q0cR4FxGGXyPuDvKCJnUtdzaAiGV4/iQLisEtKH1+0fJjn8tVJBtaOwoA4LEnx5Z3GE6dsLvr4lUn6FtI+QBxSkr0T1L9ywOahh/PGS2tvbM8ApGN1/1ob+1UTdTxd8mXPbNIHBKAWyKK2SyuRge7phB8IQE0g68y9YCUy2Dm24QdZQA2QQm51AleJleAZiYfBNiANohI3j6sTdf9pmUcxYvhA0Iby1S1LdF/q9r6hSdya/cly/WvJD0i/zVsvuLUBAAAAAElFTkSuQmCCCwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA</imageBackground>
            </state>
          </states>
        </style>
        <style role="NavigationToolbarRecentHistoryButton">
          <states>
            <state name="Normal" imageHAlign="Right" imageVAlign="Middle">
              <image>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAA1QAAAAKJUE5HDQoaCgAAAA1JSERSAAAADQAAAA4IBgAAAPR/ltIAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAA7BAAAOwQG4kWvtAAAAPklEQVQoU2P8//8/A8kApIlUTLIGsMtItWUANdmscowH4v9oeD66FzD8hKYRQwNOP0E1YtUwgAFBbJzRL3IB5VPu/4rSlhgAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</image>
            </state>
            <state name="HotTracked" imageBackgroundStyle="Centered">
              <image>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAA1QAAAAKJUE5HDQoaCgAAAA1JSERSAAAADQAAAA4IBgAAAPR/ltIAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAA7BAAAOwQG4kWvtAAAAPklEQVQoU2P8//8/A8kApIlUTLIGsMtItWUANdmscowH4v9oeD66FzD8hKYRQwNOP0E1YtUwgAFBbJzRL3IB5VPu/4rSlhgAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</image>
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAiQEAAAKJUE5HDQoaCgAAAA1JSERSAAAAFwAAABsIBgAAAJfoFNsAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAA50AAAOdAFrJLPWAAAA8klEQVRIS+2Wyw6CMBBF9eP0Q1y5YzF/q6GA4IOgUHAFbYK9YBciVdOGhQkkk9CWnCm3jzvLtm0Xkz2ATxVfwbs9WyAuabYOo8RT74QIo9hDnx4fm+BP8OR43igYcV6SEIKklMTLioIwpjg5bZHACo7ZseCgoJIU4CWaRhDG8I0VHFKU1f0NrBPlOe8ksoJDX8gwnLVu13XTrYE13ATW/TN8VPtZllmW/ip+7gTjCcVBmnfLn+0WuI/pZsSY9YLCLGBvJriTWaRptlIJOu8cJkCfzyJ7m4P5wqB9FlJR9AYN97necneD/lRa4K+cSguXgukBTWFEQnPIiugAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==</imageBackground>
            </state>
            <state name="Pressed" imageBackgroundStyle="Centered">
              <image>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAA1QAAAAKJUE5HDQoaCgAAAA1JSERSAAAADQAAAA4IBgAAAPR/ltIAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAA7BAAAOwQG4kWvtAAAAPklEQVQoU2P8//8/A8kApIlUTLIGsMtItWUANdmscowH4v9oeD66FzD8hKYRQwNOP0E1YtUwgAFBbJzRL3IB5VPu/4rSlhgAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</image>
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAiQEAAAKJUE5HDQoaCgAAAA1JSERSAAAAFwAAABsIBgAAAJfoFNsAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAA50AAAOdAFrJLPWAAAA8klEQVRIS+2Wyw6CMBBF9eP0Q1y5YzF/q6GA4IOgUHAFbYK9YBciVdOGhQkkk9CWnCm3jzvLtm0Xkz2ATxVfwbs9WyAuabYOo8RT74QIo9hDnx4fm+BP8OR43igYcV6SEIKklMTLioIwpjg5bZHACo7ZseCgoJIU4CWaRhDG8I0VHFKU1f0NrBPlOe8ksoJDX8gwnLVu13XTrYE13ATW/TN8VPtZllmW/ip+7gTjCcVBmnfLn+0WuI/pZsSY9YLCLGBvJriTWaRptlIJOu8cJkCfzyJ7m4P5wqB9FlJR9AYN97necneD/lRa4K+cSguXgukBTWFEQnPIiugAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==</imageBackground>
            </state>
          </states>
        </style>
        <style role="PrintPreviewPageNumber">
          <states>
            <state name="Normal" backColor="153, 204, 51" backGradientStyle="None" backHatchStyle="None" />
            <state name="Active" backColor="153, 204, 51" borderColor="White" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="ProgressBarFill">
          <states>
            <state name="Normal">
              <resources>
                <name>progressBarFillTrendy_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ScheduleAppointment">
          <states>
            <state name="Normal" backColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
            <state name="Selected" backColor="White" imageBackgroundStyle="Stretched" backGradientStyle="None" backHatchStyle="None" imageBackgroundStretchMargins="3, 3, 3, 3" />
          </states>
        </style>
        <style role="ScheduleCurrentDay">
          <states>
            <state name="Normal" backColor="White" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="ScheduleDay">
          <states>
            <state name="Normal">
              <resources>
                <name>scheduleDayTrendy_Normal</name>
              </resources>
            </state>
            <state name="Selected" backColor="Silver" foreColor="White" backGradientStyle="None" backHatchStyle="None" imageBackgroundStretchMargins="2, 2, 2, 2" />
            <state name="Active">
              <resources>
                <name>scheduleDayTrendy_Active</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ScheduleDayHeader">
          <states>
            <state name="Normal" backColor="Transparent" foreColor="DimGray" borderColor="Transparent" fontBold="True" backGradientStyle="None" backHatchStyle="None" imageBackgroundStretchMargins="4, 4, 4, 4" />
            <state name="Selected">
              <resources>
                <name>scheduleDayHeaderTrendy_Selected</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ScheduleDayOfWeekHeader">
          <states>
            <state name="Normal" foreColor="Gray" borderColor="Transparent" fontBold="True">
              <resources>
                <name>menuItembuttonTrendy_hotTracked</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ScheduleMonth">
          <states>
            <state name="Normal" borderColor="Gainsboro" />
          </states>
        </style>
        <style role="ScheduleMonthHeader">
          <states>
            <state name="Normal" foreColor="White" borderColor="White" imageBackgroundStyle="Stretched" fontBold="True" imageBackgroundStretchMargins="3, 3, 3, 3">
              <resources>
                <name>headerTrendy_normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ScheduleOwnerHeader">
          <states>
            <state name="Normal">
              <resources>
                <name>scheduleOwnerHeaderTrendy_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ScheduleTrailingDay">
          <states>
            <state name="Normal" foreColor="Silver" />
          </states>
        </style>
        <style role="ScheduleWeekHeader">
          <states>
            <state name="Normal" foreColor="Gray" borderColor="Transparent" backGradientStyle="Horizontal">
              <resources>
                <name>menuItembuttonTrendy_hotTracked</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ScrollBarArrowHorizontal">
          <states>
            <state name="Normal" backColor="240, 240, 240" borderColor="Transparent" backColor2="White">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAQgEAAAKJUE5HDQoaCgAAAA1JSERSAAAAFAAAABQIBgAAAI2JHQ0AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAq0lEQVQ4T63USwrCMBSFYd1t6Saapu1+fOIDFXWkiIqoiIroMprRMRE5dVxP4IM7uT93lCaAhvSF4PJVxF7poaawG3+O+wZdzdDvAY7BxbOAAoPzRw4FBmf3HAoMTm8ZFKrg1QcFGJxcMigwOD5bKDA4OlkoMDg8WigwODikUGCwv0+hwGBvZ6DAYHdroMBgZ2OgwGB7nTgPf6q+r9YqibzSQ01hN+KFYVB5AyX8h0sgLUJqAAAAAElFTkSuQmCCCwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
            </state>
          </states>
        </style>
        <style role="ScrollBarArrowVertical">
          <states>
            <state name="Normal" backColor="240, 240, 240" borderColor="Transparent" imageBackgroundStyle="Stretched" backColor2="White" backGradientStyle="Horizontal" imageBackgroundStretchMargins="5, 5, 5, 5">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAMAEAAAKJUE5HDQoaCgAAAA1JSERSAAAAFAAAABQIBgAAAI2JHQ0AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAmUlEQVQ4T2P8//8/A1UByMBDD1qzgPjHwfut/w/cbfm/73bL/723mv/vvdn8f8/15v+7rjb933ml6f+Oy03/t19s/L/1fOP/Leca/m8+2/B/0+mG/xtP1v/YcKI+C+w4qIE/gQb+p8DA/0ADfyIb+J8KBgLNQ7hw1ECSYxkUKaNhSHnCHg3DwR6GVC++MqhQwGbAiy8Qg1oYAJpDh0sROD3ZAAAAAElFTkSuQmCCCwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
            </state>
            <state name="HotTracked" foreColor="White" />
          </states>
        </style>
        <style role="ScrollBarIntersection">
          <states>
            <state name="Normal" backColor="249, 252, 252" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="ScrollBarThumbHorizontal">
          <states>
            <state name="Normal" backColor="Transparent" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAQgEAAAKJUE5HDQoaCgAAAA1JSERSAAAAFAAAABQIBgAAAI2JHQ0AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAq0lEQVQ4T63USwrCMBSFYd1t6Saapu1+fOIDFXWkiIqoiIroMprRMRE5dVxP4IM7uT93lCaAhvSF4PJVxF7poaawG3+O+wZdzdDvAY7BxbOAAoPzRw4FBmf3HAoMTm8ZFKrg1QcFGJxcMigwOD5bKDA4OlkoMDg8WigwODikUGCwv0+hwGBvZ6DAYHdroMBgZ2OgwGB7nTgPf6q+r9YqibzSQ01hN+KFYVB5AyX8h0sgLUJqAAAAAElFTkSuQmCCCwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
            </state>
          </states>
        </style>
        <style role="ScrollBarThumbVertical">
          <states>
            <state name="Normal" backColor="Transparent" borderColor="Transparent" imageBackgroundStyle="Stretched" backGradientStyle="None" backHatchStyle="None">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAMAEAAAKJUE5HDQoaCgAAAA1JSERSAAAAFAAAABQIBgAAAI2JHQ0AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAmUlEQVQ4T2P8//8/A1UByMBDD1qzgPjHwfut/w/cbfm/73bL/723mv/vvdn8f8/15v+7rjb933ml6f+Oy03/t19s/L/1fOP/Leca/m8+2/B/0+mG/xtP1v/YcKI+C+w4qIE/gQb+p8DA/0ADfyIb+J8KBgLNQ7hw1ECSYxkUKaNhSHnCHg3DwR6GVC++MqhQwGbAiy8Qg1oYAJpDh0sROD3ZAAAAAElFTkSuQmCCCwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
            </state>
          </states>
        </style>
        <style role="ScrollBarTrackHorizontal">
          <states>
            <state name="Normal" backColor="240, 240, 240" backColor2="White" backGradientStyle="Vertical" />
          </states>
        </style>
        <style role="ScrollBarTrackVertical">
          <states>
            <state name="Normal" backColor="240, 240, 240" backColor2="White" backGradientStyle="Horizontal" />
          </states>
        </style>
        <style role="SpinButton">
          <states>
            <state name="Normal" borderColor="Transparent" foreColorDisabled="152, 194, 143" />
          </states>
        </style>
        <style role="StateEditorButton">
          <states>
            <state name="Normal" backColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="StatusBarPanel">
          <states>
            <state name="Normal" borderColor="Transparent" />
          </states>
        </style>
        <style role="StatusBarPanelStateButton">
          <states>
            <state name="Normal" backColor="249, 252, 252" foreColor="White" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="StatusBarProgressBar">
          <states>
            <state name="Normal" backColor="249, 252, 252" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="TabControlClientArea">
          <states>
            <state name="Normal" backColor="249, 252, 252" borderColor="227, 241, 199" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="TabControlTabsArea">
          <states>
            <state name="Normal" backColor="249, 252, 252" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="TabControlTabsAreaHorizontalTop">
          <states>
            <state name="Normal" backColor="249, 252, 252" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="TabItemArea">
          <states>
            <state name="Normal" backColor="249, 252, 252" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="TabItemHorizontalBottom">
          <states>
            <state name="Normal">
              <resources>
                <name>tabItemHorizontalBottomTrendy_Normal</name>
              </resources>
            </state>
            <state name="Selected">
              <resources>
                <name>tabItemHorizontalBottomTrendy_Selected</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="TabItemHorizontalTop">
          <states>
            <state name="Normal">
              <resources>
                <name>tabItemHorizontalTopTrendy_Normal</name>
              </resources>
            </state>
            <state name="Selected">
              <resources>
                <name>tabItemHorizontalTopTrendy_Selected</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="TabItemVerticalLeft">
          <states>
            <state name="Normal">
              <resources>
                <name>tabItemVerticalLeftTrendy_Normal</name>
              </resources>
            </state>
            <state name="Selected">
              <resources>
                <name>tabItemVerticalLeftTrendy_Selected</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="TabItemVerticalRight">
          <states>
            <state name="Normal">
              <resources>
                <name>tabItemVerticalRightTrendy_Normal</name>
              </resources>
            </state>
            <state name="Selected">
              <resources>
                <name>tabItemVerticalRightTrendy_Selected</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="TaskPane">
          <states>
            <state name="Normal" backColor="White" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="TaskPaneToolbar">
          <states>
            <state name="Normal">
              <resources>
                <name>headerTrendy_normal</name>
              </resources>
            </state>
            <state name="Active">
              <resources>
                <name>headerTrendy_normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="Toolbar">
          <states>
            <state name="Normal" backColor="249, 252, 252" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="ToolbarCloseButton">
          <states>
            <state name="Normal" backColor="Transparent" foreColor="White" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
            <state name="HotTracked" backColor="Transparent" foreColor="Black" imageBackgroundStyle="Stretched" backGradientStyle="None" backHatchStyle="None">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAtAAAAAKJUE5HDQoaCgAAAA1JSERSAAAACgAAAAoIBgAAAI0yz70AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAHUlEQVQoU2P8//8/A1EApJAYTJQisK3EmDZyFQIALPcQACGanOcAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
            </state>
          </states>
        </style>
        <style role="ToolbarDockArea">
          <states>
            <state name="Normal" backColor="250, 252, 253" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="ToolbarFloatingCaption">
          <states>
            <state name="Normal" backColor="153, 204, 51" foreColor="White" borderColor="White" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="ToolbarItem">
          <states>
            <state name="Normal" backColor="250, 252, 253" backGradientStyle="None" backHatchStyle="None" />
            <state name="HotTracked" foreColor="WhiteSmoke" borderColor="Transparent" imageBackgroundStretchMargins="3, 3, 3, 3">
              <resources>
                <name>toobarItemPopupMenuTrendy_HotTracked</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ToolbarItemPopup">
          <states>
            <state name="HotTracked" backColor="249, 252, 252" foreColor="WhiteSmoke" borderColor="Gainsboro" backGradientStyle="None" backHatchStyle="None" imageBackgroundStretchMargins="3, 3, 3, 3" />
          </states>
        </style>
        <style role="ToolbarItemPopupColorPicker">
          <states>
            <state name="HotTracked" borderColor="Transparent" />
          </states>
        </style>
        <style role="ToolbarItemPopupMenu">
          <states>
            <state name="Normal" backColor="250, 252, 253" backGradientStyle="None" backHatchStyle="None" />
            <state name="HotTracked">
              <resources>
                <name>toobarItemPopupMenuTrendy_HotTracked</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ToolbarItemProgressBar">
          <states>
            <state name="Normal" borderColor="Gainsboro" />
          </states>
        </style>
        <style role="ToolbarItemQuickCustomize">
          <states>
            <state name="HotTracked" foreColor="SlateGray">
              <resources>
                <name>shareTransparentImageTrendy</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ToolbarItemTaskPaneLabel" borderStyle="None">
          <states>
            <state name="Normal" backColor="Transparent" foreColor="White" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="ToolbarItemTaskPaneMenuDropDownOnly">
          <states>
            <state name="Normal" backColor="Transparent" foreColor="White" borderColor="Transparent" imageBackgroundStyle="Stretched" backGradientStyle="None" backHatchStyle="None">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAtAAAAAKJUE5HDQoaCgAAAA1JSERSAAAACgAAAAoIBgAAAI0yz70AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAHUlEQVQoU2P8//8/A1EApJAYTJQisK3EmDZyFQIALPcQACGanOcAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
            </state>
            <state name="HotTracked" backColor="Transparent" foreColor="Black" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAtAAAAAKJUE5HDQoaCgAAAA1JSERSAAAACgAAAAoIBgAAAI0yz70AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAHUlEQVQoU2P8//8/A1EApJAYTJQisK3EmDZyFQIALPcQACGanOcAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
            </state>
            <state name="Pressed" backColor="Transparent" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
            <state name="Active" backColor="Transparent" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="ToolbarItemTaskPaneNavigation">
          <states>
            <state name="Normal" backColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
            <state name="HotTracked" backColor="Transparent" backGradientStyle="None" backHatchStyle="None">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAtAAAAAKJUE5HDQoaCgAAAA1JSERSAAAACgAAAAoIBgAAAI0yz70AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAHUlEQVQoU2P8//8/A1EApJAYTJQisK3EmDZyFQIALPcQACGanOcAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
            </state>
          </states>
        </style>
        <style role="ToolTip">
          <states>
            <state name="Normal" foreColor="White" borderColor="Transparent" backGradientStyle="Elliptical">
              <resources>
                <name>gridCellTrendy_hotTracked</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ToolTipBalloon">
          <states>
            <state name="Normal" foreColor="White" borderColor="DimGray" backGradientStyle="Elliptical">
              <resources>
                <name>gridCellTrendy_hotTracked</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="ToolTipBalloonTitle">
          <states>
            <state name="Normal" foreColor="White" />
          </states>
        </style>
        <style role="TreeCell" borderStyle="None">
          <states>
            <state name="Normal" backColor="249, 252, 252" backGradientStyle="None" backHatchStyle="None" />
            <state name="Active" foreColor="White" />
          </states>
        </style>
        <style role="TreeColumnHeader" borderStyle="None">
          <states>
            <state name="Normal" foreColor="White" borderColor="Transparent" fontBold="True">
              <cursor>Hand</cursor>
              <resources>
                <name>headerTrendy_normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="TreeControlArea">
          <states>
            <state name="Normal" backColor="249, 252, 252" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="TreeNode">
          <states>
            <state name="Selected" backColor="0, 153, 255" borderColor="White" imageBackgroundStyle="Stretched" backGradientStyle="None" backHatchStyle="None" imageBackgroundStretchMargins="3, 3, 3, 3" />
            <state name="Active" imageBackgroundStyle="Stretched" imageBackgroundStretchMargins="3, 3, 3, 3">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAA1gAAAAKJUE5HDQoaCgAAAA1JSERSAAAAFAAAABQIBgAAAI2JHQ0AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAP0lEQVQ4T2P8DwQM1AQgA6kJGGCGrTjx+j8lGGbOqIHkh+NoGP6nKA2C0u9oGI6GIRlFGUayoVYhCy8PqWUgAA6aTFlyxLyuAAAAAElFTkSuQmCCCwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
            </state>
          </states>
        </style>
        <style role="TreeNodeExpansionIndicator">
          <states>
            <state name="Normal" borderColor="White" />
          </states>
        </style>
        <style role="UltraCalculator">
          <states>
            <state name="Normal" borderColor="White">
              <resources>
                <name>ExplorerBarGroupItemAreaInnerTrendy_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UltraCalculatorButtonBase">
          <states>
            <state name="Normal" backColor="Transparent" backGradientStyle="None" backHatchStyle="None">
              <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAIQEAAAKJUE5HDQoaCgAAAA1JSERSAAAAFAAAABQIBgAAAI2JHQ0AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAiklEQVQ4T2P8//8/A1UByMBtt0uyJp12+lG6T/J/wR4hkjBID0gvyAyw40BE30mbn6QahK4eZAbcQHJchm4gyAy4gZS6DqZ/1EDSkgu2cB8Nw9EwJLHUASUjeLIp2Ut6sYWeDkFmwA1sP2BFcfEFMgNu4PJjBRndBx1+kONSkB6QXpAZcANBDGphAIiHgu4SgCabAAAAAElFTkSuQmCCCwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
            </state>
            <state name="HotTracked" foreColor="Gray" borderColor="Transparent" />
          </states>
        </style>
        <style role="UltraCalculatorEditArea">
          <states>
            <state name="Normal" backColor="White" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="UltraComboEditPortion">
          <states>
            <state name="Normal" foreColor="DimGray" borderColor="LightGray" />
          </states>
        </style>
        <style role="UltraGroupBox">
          <states>
            <state name="Normal" backColor="249, 252, 252" borderColor="249, 252, 252" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="UltraGroupBoxContentArea">
          <states>
            <state name="Normal">
              <resources>
                <name>ultraGroupBoxContentAreaTrendy_Normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UltraGroupBoxHeaderHorizontalTop">
          <states>
            <state name="Normal" backColor="Transparent" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="UltraOptionSet">
          <states>
            <state name="Normal" backColor="249, 252, 252" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="UltraPictureBox">
          <states>
            <state name="Normal">
              <resources>
                <name>controlAreaTrendy_normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UltraProgressBar">
          <states>
            <state name="Normal" backColor="249, 252, 252" borderColor="Gainsboro" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="UltraStatusBar">
          <states>
            <state name="Normal" backColor="249, 252, 252" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="UnpinnedTabAreaBottom">
          <states>
            <state name="Normal" backColor="249, 252, 252" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="UnpinnedTabAreaLeft">
          <states>
            <state name="Normal" backColor="249, 252, 252" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="UnpinnedTabAreaRight">
          <states>
            <state name="Normal" backColor="249, 252, 252" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="UnpinnedTabAreaTop">
          <states>
            <state name="Normal">
              <resources>
                <name>headerTrendy_normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UnpinnedTabItemHorizontal">
          <states>
            <state name="Normal" foreColor="White" fontBold="True">
              <resources>
                <name>menuItembuttonTrendy_hotTracked</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UnpinnedTabItemHorizontalBottom" buttonStyle="FlatBorderless">
          <states>
            <state name="Normal" borderColor="Transparent">
              <resources>
                <name>headerTrendy_normal</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UnpinnedTabItemVertical">
          <states>
            <state name="Normal" foreColor="White" borderColor="White" backGradientStyle="Horizontal">
              <resources>
                <name>menuItembuttonTrendy_hotTracked</name>
              </resources>
            </state>
          </states>
        </style>
        <style role="UnpinnedTabItemVerticalRight">
          <states>
            <state name="Normal" backColor="249, 252, 252" backGradientStyle="None" backHatchStyle="None" />
          </states>
        </style>
        <style role="ValueList">
          <states>
            <state name="Normal" borderColor="Gainsboro" />
          </states>
        </style>
        <style role="ValueListItem">
          <states>
            <state name="Normal" foreColor="Gray" borderColor="Gainsboro" />
            <state name="Selected" backColor="DimGray" foreColor="White" imageBackgroundStyle="Stretched" backGradientStyle="None" backHatchStyle="None" imageBackgroundStretchMargins="3, 3, 3, 3" />
          </states>
        </style>
      </styles>
      <sharedObjects>
        <sharedObject name="ScrollBar">
          <properties>
            <property name="MinimumThumbHeight" colorCategory="{Default}">20</property>
          </properties>
        </sharedObject>
      </sharedObjects>
    </styleSet>
  </styleSets>
  <resources>
    <resource name="buttonTrendy_Normal" foreColor="White" borderColor="White" />
    <resource name="controlAreaTrendy_normal" backColor="244, 255, 229" borderColor="Gainsboro" backColor2="249, 252, 252" backGradientStyle="Vertical" />
    <resource name="dayViewTimeSlotDescriptorTrendy_Normal" backColor="249, 252, 252" borderColor="Gainsboro" fontBold="True" backGradientStyle="None" backHatchStyle="None" />
    <resource name="dropDownButtonTrendy_Normal" backColor="Transparent" foreColor="White" borderColor="White" imageBackgroundStyle="Stretched" backColorAlpha="Opaque" backGradientStyle="None" backHatchStyle="None" imageBackgroundStretchMargins="6, 6, 6, 6">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAzQEAAAKJUE5HDQoaCgAAAA1JSERSAAAAFAAAABQIBgAAAI2JHQ0AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAABNklEQVQ4T62U2UrDUBCGR7o8gigIxZv4Itr2YnxAEQRRUEERRChVBBdcERdU0KcZv2HS0HNTPSGBj1ky/88hmWROnsyk0csNm0QeMCwp3syG32brP2ZK/AufXUEz7SFySwOKF7PVT7N5Ttu6i94s2swsMNv/Cm01K9cUMPjA7D7yHJYeQ1tp5JIClGbrKvIcOmhcW2nkggL0PWIdEq2cYQL6GrEOiVZGmIA+R6xDopVTTEB9F8s8NyZaOcEIlLc1yXNjopVjjEB9Zco8NyZaOcQI1Be5zHNjopUDjEBvItYh0coeJqAsaHs/8hy6fhj/ICY62aWAAUu9eBR5Dj2eu2srjWxTQMGNNRa7x3Ps7ERvFl1mljlA/zy01axsUZQUmA3H/L4Y+g8+65ppD5FNGk0iGxg2yC8FYBZOkXTOIgAAAABJRU5ErkJgggsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="editorButtonTrendy_Normal" backColor="Transparent" imageBackgroundStyle="Stretched" backGradientStyle="None" backHatchStyle="None" imageBackgroundStretchMargins="5, 5, 5, 5">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAQgEAAAKJUE5HDQoaCgAAAA1JSERSAAAAFAAAABQIBgAAAI2JHQ0AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAq0lEQVQ4T63USwrCMBSFYd1t6Saapu1+fOIDFXWkiIqoiIroMprRMRE5dVxP4IM7uT93lCaAhvSF4PJVxF7poaawG3+O+wZdzdDvAY7BxbOAAoPzRw4FBmf3HAoMTm8ZFKrg1QcFGJxcMigwOD5bKDA4OlkoMDg8WigwODikUGCwv0+hwGBvZ6DAYHdroMBgZ2OgwGB7nTgPf6q+r9YqibzSQ01hN+KFYVB5AyX8h0sgLUJqAAAAAElFTkSuQmCCCwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="ExplorerBarGroupItemAreaInnerTrendy_Normal" backColor="225, 248, 192" backColor2="179, 217, 113" backGradientStyle="Vertical" />
    <resource name="gridAddNewBoxTrendy_Normal" imageBackgroundStyle="Stretched" imageBackgroundStretchMargins="4, 4, 4, 4">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAzQEAAAKJUE5HDQoaCgAAAA1JSERSAAAAFAAAABQIBgAAAI2JHQ0AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAABNklEQVQ4T62U2UrDUBCGR7o8gigIxZv4Itr2YnxAEQRRUEERRChVBBdcERdU0KcZv2HS0HNTPSGBj1ky/88hmWROnsyk0csNm0QeMCwp3syG32brP2ZK/AufXUEz7SFySwOKF7PVT7N5Ttu6i94s2swsMNv/Cm01K9cUMPjA7D7yHJYeQ1tp5JIClGbrKvIcOmhcW2nkggL0PWIdEq2cYQL6GrEOiVZGmIA+R6xDopVTTEB9F8s8NyZaOcEIlLc1yXNjopVjjEB9Zco8NyZaOcQI1Be5zHNjopUDjEBvItYh0coeJqAsaHs/8hy6fhj/ICY62aWAAUu9eBR5Dj2eu2srjWxTQMGNNRa7x3Ps7ERvFl1mljlA/zy01axsUZQUmA3H/L4Y+g8+65ppD5FNGk0iGxg2yC8FYBZOkXTOIgAAAABJRU5ErkJgggsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="gridCardCaptionTrendy_Normal" backColor="LightSkyBlue" backColor2="0, 153, 255" backHatchStyle="LightUpwardDiagonal" />
    <resource name="gridCellTrendy_EditMode" backColor="White" foreColor="DimGray" borderColor="Transparent" imageBackgroundStyle="Stretched" backColorAlpha="Opaque" foregroundAlpha="Opaque" backGradientStyle="None" backHatchStyle="None" imageBackgroundStretchMargins="3, 3, 3, 3" />
    <resource name="gridCellTrendy_hotTracked" backColor="DeepSkyBlue" foreColor="Azure" borderColor="Transparent" imageBackgroundStyle="Stretched" backColor2="51, 204, 255" backGradientStyle="Vertical" imageBackgroundStretchMargins="3, 3, 3, 3" />
    <resource name="gridExpansionIndicatorTrendy_Normal" backColor="Transparent" foreColor="Gainsboro" borderColor="LightGray" backGradientStyle="None" backHatchStyle="None" />
    <resource name="gridRowSelectorTrendy_Normal" backColor="White" foreColor="Silver" borderColor="White" backGradientStyle="None" backHatchStyle="None" />
    <resource name="gridRowTrendy_Active" backColor="DeepSkyBlue" foreColor="White" borderColor="Gainsboro" fontSize="8" backColor2="0, 153, 255" backHatchStyle="DarkUpwardDiagonal" />
    <resource name="gridRowTrendy_Selected" backColor="235, 248, 255" foreColor="DarkGray" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None" />
    <resource name="headerTrendy_normal" backColor="153, 204, 51" foreColor="White" borderColor="White" imageBackgroundStyle="Stretched" fontBold="True" fontSize="9" backColor2="0, 149, 74" backGradientStyle="Vertical" imageBackgroundStretchMargins="6, 6, 6, 6" />
    <resource name="listViewItemTrendy_HotTracked" imageBackgroundStyle="Stretched" imageBackgroundStretchMargins="3, 3, 3, 3">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAA1gAAAAKJUE5HDQoaCgAAAA1JSERSAAAAFAAAABQIBgAAAI2JHQ0AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAP0lEQVQ4T2P8DwQM1AQgA6kJGGCGrTjx+j8lGGbOqIHkh+NoGP6nKA2C0u9oGI6GIRlFGUayoVYhCy8PqWUgAA6aTFlyxLyuAAAAAElFTkSuQmCCCwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="mainMenubarTrendy_Normal" backColor="250, 252, 253" backGradientStyle="None" backHatchStyle="None" />
    <resource name="menuItembuttonTrendy_hotTracked" backColor="WhiteSmoke" borderColor="White" backGradientStyle="None" backHatchStyle="None" />
    <resource name="progressBarFillTrendy_Normal" backColor="DeepSkyBlue" foreColor="White" fontBold="True" backColor2="DodgerBlue" backGradientStyle="Vertical" imageBackgroundStretchMargins="3, 3, 3, 3" />
    <resource name="scheduleDayHeaderTrendy_Selected" backColor="153, 204, 51" foreColor="White" borderColor="White" imageBackgroundStyle="Stretched" fontBold="True" backColor2="0, 149, 74" backGradientStyle="Vertical" imageBackgroundStretchMargins="6, 6, 6, 6" />
    <resource name="scheduleDayTrendy_Active" backColor="0, 153, 255" foreColor="White" borderColor="Transparent" imageBackgroundStyle="Stretched" backColor2="0, 204, 255" backGradientStyle="Vertical" imageBackgroundStretchMargins="2, 2, 2, 2" />
    <resource name="scheduleDayTrendy_Normal" backColor="White" foreColor="DimGray" borderColor="Gainsboro" imageBackgroundStyle="Stretched" backColor2="240, 240, 240" backGradientStyle="ForwardDiagonal" imageBackgroundStretchMargins="0, 0, 0, 15" />
    <resource name="scheduleOwnerHeaderTrendy_Normal" backColor="153, 204, 51" foreColor="White" borderColor="Transparent" imageBackgroundStyle="Stretched" fontSize="9" backColor2="0, 149, 74" backGradientStyle="Vertical" imageBackgroundStretchMargins="4, 4, 4, 4" />
    <resource name="shareTransparentImageTrendy" backColor="Transparent" backGradientStyle="None" backHatchStyle="None">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAtAAAAAKJUE5HDQoaCgAAAA1JSERSAAAACgAAAAoIBgAAAI0yz70AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAHUlEQVQoU2P8//8/A1EApJAYTJQisK3EmDZyFQIALPcQACGanOcAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="spacer" backColor="Transparent" borderColor="Transparent" backGradientStyle="None" backHatchStyle="None">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAqQAAAAKJUE5HDQoaCgAAAA1JSERSAAAAAgAAAAIIBgAAAHK2DSQAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAEklEQVQYV2P8//8/AxiAGCAMAE3WCPnXj6AWAAAAAElFTkSuQmCCCwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="tabItemHorizontalBottomTrendy_Normal" backColor="Transparent" foreColor="211, 234, 181" borderColor="Transparent" textVAlign="Top" imageBackgroundStyle="Stretched" imageBackgroundAlpha="Opaque" foregroundAlpha="Opaque" backGradientStyle="None" backHatchStyle="None" imageBackgroundStretchMargins="3, 0, 3, 5">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAACgEAAAKJUE5HDQoaCgAAAA1JSERSAAAAEgAAAA0IBgAAAKROPhkAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAArvAAAK7wF9dopIAAAAc0lEQVQ4T2OsP6T/n4EaoOWI6X9qYIbOYzb/qYEZ+k46/acGZph82us/NTDDzLMB/6mBGeadD/tPDcyw+FLsf2pghhVXU/5TAzOsu571nxqYYfPNwv/UwAw77lTsB+L/FOL9DHvv1fMD8Xog/k8mBunlBwCNE1o1eGy24gAAAABJRU5ErkJgggsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="tabItemHorizontalBottomTrendy_Selected" foreColor="White" textVAlign="Top" imageBackgroundStyle="Stretched" fontBold="True" imageBackgroundStretchMargins="3, 0, 3, 5">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAACQEAAAKJUE5HDQoaCgAAAA1JSERSAAAAEgAAAA0IBgAAAKROPhkAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAArvAAAK7wF9dopIAAAAcklEQVQ4T2NkmMnwn4EagHce539qYAaZxSL/qYEZtFfK/acGZrBep/GfGpjBa4vxf2pghshdtv+pgRky9rv/pwZmKD8S+J8amKHtVMx/amCGqRfS/lMDMyy5WrAfiP9TiPczbLldzQ/E64H4P5kYpJcfAFo4/Z9RfnX0AAAAAElFTkSuQmCCCwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="tabItemHorizontalTopTrendy_Normal" backColor="249, 252, 252" foreColor="211, 234, 181" borderColor="Transparent" textHAlign="Left" textVAlign="Bottom" imageBackgroundStyle="Stretched" imageBackgroundAlpha="Opaque" foregroundAlpha="Opaque" backGradientStyle="None" backHatchStyle="None" imageBackgroundStretchMargins="3, 5, 3, 0">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAADgEAAAKJUE5HDQoaCgAAAA1JSERSAAAAEgAAAA0IBgAAAKROPhkAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAArvAAAK7wF9dopIAAAAd0lEQVQ4T2M89KBVgIGBYQEQ+wMxOWAjUFMCw9579fuB+D+FeD/DjjsV/6mBGTbfLPxPDcyw7nrWf2pghhVXU/5TAzMsvhT7nxqYYd75sP/UwAwzzwb8pwZmmHza6z81MEPfSaf/1MAMncds/lMDM7QcMf1PDQwAuGVefYwnzlMAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="tabItemHorizontalTopTrendy_Selected" backColor="Transparent" foreColor="White" borderColor="Transparent" textHAlign="Left" imageBackgroundStyle="Stretched" fontBold="True" backGradientStyle="None" backHatchStyle="None" imageBackgroundStretchMargins="3, 5, 3, 0">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAADQEAAAKJUE5HDQoaCgAAAA1JSERSAAAAEgAAAA0IBgAAAKROPhkAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAArvAAAK7wF9dopIAAAAdklEQVQ4T2M89KBVgIGBYQEQ+wMxOWAjUFMCw5bb1fuB+D+FeD/DkqsF/6mBGaZeSPtPDczQdirmPzUwQ/mRwP/UwAwZ+93/UwMzRO6y/U8NzOC1xfg/NTCD9TqN/9TADNor5f5TAzPILBb5Tw3MwDuP8z81MABfCALN5O6oKAAAAABJRU5ErkJgggsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="tabItemVerticalLeftTrendy_Normal" backColor="Transparent" foreColor="White" borderColor="Transparent" textHAlign="Right" textVAlign="Middle" imageBackgroundStyle="Stretched" imageBackgroundAlpha="Opaque" foregroundAlpha="Opaque" fontBold="True" backGradientStyle="None" backHatchStyle="None" imageBackgroundStretchMargins="5, 3, 0, 3">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAFAEAAAKJUE5HDQoaCgAAAA1JSERSAAAADQAAABIIBgAAAIBrVDIAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAArvAAAK7wF9dopIAAAAfUlEQVQ4T2M89KBV4Pe/Xwt+//vp//vvTwYgzfDr3y8GMBvM/wXh//sB5IPonwwMQE37996r/7/jTsX/zTcL/6+7nvV/xdWU/4svxf6fdz7s/8yzAf8nn/b633fS6X/nMZv/LUdM/4M0/R/VNBoQwz5FkJU1+IFZYz0p+QkA0N+TjY2kmrgAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="tabItemVerticalLeftTrendy_Selected" imageBackgroundStyle="Stretched" imageBackgroundStretchMargins="5, 3, 0, 3">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAADgEAAAKJUE5HDQoaCgAAAA1JSERSAAAADQAAABIIBgAAAIBrVDIAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAArvAAAK7wF9dopIAAAAd0lEQVQ4T2M89KBV4NPvbws+/Prq//H3N4aPv0D4KwOYjcT/AGP//srAANS0f8vt6v9Lrhb8n3oh7X/bqZj/5UcC/2fsd/8fucv2v9cW4//W6zT+a6+U+y+zWOQ/7zzO/yBN/0c1jQbEsE8RZGUNfmDWWE9KfgIAafg8LP7bNFUAAAAASUVORK5CYIILAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="tabItemVerticalRightTrendy_Normal" backColor="Transparent" foreColor="211, 234, 181" borderColor="Transparent" textHAlign="Left" textVAlign="Bottom" imageBackgroundStyle="Stretched" imageBackgroundAlpha="Opaque" foregroundAlpha="Opaque" backGradientStyle="None" backHatchStyle="None" imageBackgroundStretchMargins="0, 3, 5, 3">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAEAEAAAKJUE5HDQoaCgAAAA1JSERSAAAADQAAABIIBgAAAIBrVDIAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAArvAAAK7wF9dopIAAAAeUlEQVQ4T2OsP6T/n5WJjYGViZ2BlRlEczCwgflAzAwSY4fyweyNQHUJDCBNLUdM/3ces/nfd9Lp/+TTXv9nng34P+982P/Fl2L/r7ia8n/d9az/m28W/t9xp+L/3nv1+8nR9H9UEyT0RgMCmoyGY0CQnDXWA1MEPwAHiGUUUNssOQAAAABJRU5ErkJgggsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="tabItemVerticalRightTrendy_Selected" backColor="Transparent" foreColor="White" borderColor="Transparent" textHAlign="Left" textVAlign="Bottom" imageBackgroundStyle="Stretched" fontBold="True" backGradientStyle="None" backHatchStyle="None" imageBackgroundStretchMargins="0, 3, 5, 3">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAADQEAAAKJUE5HDQoaCgAAAA1JSERSAAAADQAAABIIBgAAAIBrVDIAAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAArvAAAK7wF9dopIAAAAdklEQVQ4T2NkmMnwn5eVk4GflZuBn42LQYCVC0hD2PwgNip/owAbdwIDWNM8zv8yi0X+a6+U+2+9TuO/1xbj/5G7bP9n7Hf/X34k8H/bqZj/Uy+k/V9yteD/ltvV+8nR9H9UEyT0RgMCmoyGY0CQnDXWA1MEPwCvCAAl2wcXwgAAAABJRU5ErkJgggsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="toobarItemPopupMenuTrendy_HotTracked" foreColor="WhiteSmoke" borderColor="Transparent" imageBackgroundStyle="Stretched" imageBackgroundStretchMargins="10, 9, 10, 9">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAAzQEAAAKJUE5HDQoaCgAAAA1JSERSAAAAFAAAABQIBgAAAI2JHQ0AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAABNklEQVQ4T62U2UrDUBCGR7o8gigIxZv4Itr2YnxAEQRRUEERRChVBBdcERdU0KcZv2HS0HNTPSGBj1ky/88hmWROnsyk0csNm0QeMCwp3syG32brP2ZK/AufXUEz7SFySwOKF7PVT7N5Ttu6i94s2swsMNv/Cm01K9cUMPjA7D7yHJYeQ1tp5JIClGbrKvIcOmhcW2nkggL0PWIdEq2cYQL6GrEOiVZGmIA+R6xDopVTTEB9F8s8NyZaOcEIlLc1yXNjopVjjEB9Zco8NyZaOcQI1Be5zHNjopUDjEBvItYh0coeJqAsaHs/8hy6fhj/ICY62aWAAUu9eBR5Dj2eu2srjWxTQMGNNRa7x3Ps7ERvFl1mljlA/zy01axsUZQUmA3H/L4Y+g8+65ppD5FNGk0iGxg2yC8FYBZOkXTOIgAAAABJRU5ErkJgggsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="ultraCalculatorButtonActionTrendy_Normal" foreColor="Firebrick" borderColor="Gainsboro" imageBackgroundStyle="Stretched" fontBold="True" imageBackgroundStretchMargins="4, 4, 4, 4">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAA1gAAAAKJUE5HDQoaCgAAAA1JSERSAAAAFAAAABQIBgAAAI2JHQ0AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAP0lEQVQ4T2P8DwQM1AQgA6kJGGCGvT694j8lGGbOqIHkh+NoGP6nKA2C0u9oGI6GIRlFGUayoVYhCy8PqWUgABSwT1n0SVNAAAAAAElFTkSuQmCCCwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="ultraCalculatorButtonImmediateActionTrendy_Normal">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAA1gAAAAKJUE5HDQoaCgAAAA1JSERSAAAAFAAAABQIBgAAAI2JHQ0AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAP0lEQVQ4T2P8DwQM1AQgA6kJGGCGvT694j8lGGbOqIHkh+NoGP6nKA2C0u9oGI6GIRlFGUayoVYhCy8PqWUgABSwT1n0SVNAAAAAAElFTkSuQmCCCwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="ultraCalculatorButtonNumericTrendy_Normal" borderColor="Gainsboro" imageBackgroundStyle="Stretched" imageBackgroundStretchMargins="3, 3, 3, 3">
      <imageBackground>AAEAAAD/////AQAAAAAAAAAMAgAAAFFTeXN0ZW0uRHJhd2luZywgVmVyc2lvbj0yLjAuMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWIwM2Y1ZjdmMTFkNTBhM2EFAQAAABVTeXN0ZW0uRHJhd2luZy5CaXRtYXABAAAABERhdGEHAgIAAAAJAwAAAA8DAAAA1gAAAAKJUE5HDQoaCgAAAA1JSERSAAAAFAAAABQIBgAAAI2JHQ0AAAABc1JHQgCuzhzpAAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAACXBIWXMAAAsQAAALEAGtI711AAAAP0lEQVQ4T2P8DwQM1AQgA6kJGGCGvT694j8lGGbOqIHkh+NoGP6nKA2C0u9oGI6GIRlFGUayoVYhCy8PqWUgABSwT1n0SVNAAAAAAElFTkSuQmCCCwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=</imageBackground>
    </resource>
    <resource name="ultraGroupBoxContentAreaTrendy_Normal" backColor="249, 252, 252" backColor2="249, 252, 252" backGradientStyle="ForwardDiagonal" />
  </resources>
</styleLibrary>