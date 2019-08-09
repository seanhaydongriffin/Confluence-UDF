#include-once
#Include <Array.au3>
#include <GuiEdit.au3>
#include "cURL.au3"
#include <SQLite.au3>
#include <SQLite.dll.au3>
#include <Toast.au3>
#include <Crypt.au3>
#include <Jira.au3>
#include <Date.au3>
#Region Header
#cs
	Title:   		Janison Insights Automation UDF Library for AutoIt3
	Filename:  		JanisonInsights.au3
	Description: 	A collection of functions for creating, attaching to, reading from and manipulating Janison Insights
	Author:   		seangriffin
	Version:  		V0.1
	Last Update: 	25/02/18
	Requirements: 	AutoIt3 3.2 or higher,
					Janison Insights Release x.xx,
					cURL xxx
	Changelog:		---------24/12/08---------- v0.1
					Initial release.
#ce
#EndRegion Header
#Region Global Variables and Constants
Global $confluence_domain = ""
Global $confluence_username = ""
Global $confluence_password = ""
Global $confluence_encrypted_password = ""
Global $confluence_decrypted_password = ""
Global $confluence_json = ""
Global $confluence_html = ""
Global $storage_format
Global $aResult, $iRows, $iColumns, $iRval
#EndRegion Global Variables and Constants
#Region Core functions
; #FUNCTION# ;===============================================================================
;
; Name...........:	_InsightsSetup()
; Description ...:	Setup activities including cURL initialization.
; Syntax.........:	_InsightsSetup()
; Parameters ....:
; Return values .: 	On Success			- Returns True.
;                 	On Failure			- Returns False, and:
;											sets @ERROR = 1 if unable to find an active SAP session.
;												This means the SAP GUI Scripting interface is not enabled.
;												Refer to the "Requirements" section at the top of this file.
;											sets @ERROR = 2 if unable to find the SAP window to attach to.
;
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:	A prerequisite is that the SAP GUI Scripting interface is enabled,
;					and the SAP user is already logged in (ie. The "SAP Easy Access" window is displayed).
;					Refer to the "Requirements" section at the top of this file for information
;					on enabling the SAP GUI Scripting interface.
;
; Related .......:
; Link ..........:
; Example .......:	Yes
;
; ;==========================================================================================
Func _ConfluenceSetup()

	; Initialise cURL
	cURL_initialise()


EndFunc

; #FUNCTION# ;===============================================================================
;
; Name...........:	_InsightsShutdown()
; Description ...:	Setup activities including cURL initialization.
; Syntax.........:	_InsightsShutdown()
; Parameters ....:
; Return values .: 	On Success			- Returns True.
;                 	On Failure			- Returns False, and:
;											sets @ERROR = 1 if unable to find an active SAP session.
;												This means the SAP GUI Scripting interface is not enabled.
;												Refer to the "Requirements" section at the top of this file.
;											sets @ERROR = 2 if unable to find the SAP window to attach to.
;
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:	A prerequisite is that the SAP GUI Scripting interface is enabled,
;					and the SAP user is already logged in (ie. The "SAP Easy Access" window is displayed).
;					Refer to the "Requirements" section at the top of this file for information
;					on enabling the SAP GUI Scripting interface.
;
; Related .......:
; Link ..........:
; Example .......:	Yes
;
; ;==========================================================================================
Func _ConfluenceShutdown()

	; Clean up cURL
	cURL_cleanup()

EndFunc


; #FUNCTION# ;===============================================================================
;
; Name...........:	_InsightsDomainSet()
; Description ...:	Sets the domain to use in all other functions.
; Syntax.........:	_InsightsDomainSet($domain)
; Parameters ....:	$win_title			- Optional: The title of the SAP window (within the session) to attach to.
;											The window "SAP Easy Access" is used if one isn't provided.
;											This may be a substring of the full window title.
;					$sap_transaction	- Optional: a SAP transaction to run after attaching to the session.
;											A "/n" will be inserted at the beginning of the transaction
;											if one isn't provided.
; Return values .: 	On Success			- Returns True.
;                 	On Failure			- Returns False, and:
;											sets @ERROR = 1 if unable to find an active SAP session.
;												This means the SAP GUI Scripting interface is not enabled.
;												Refer to the "Requirements" section at the top of this file.
;											sets @ERROR = 2 if unable to find the SAP window to attach to.
;
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:	A prerequisite is that the SAP GUI Scripting interface is enabled,
;					and the SAP user is already logged in (ie. The "SAP Easy Access" window is displayed).
;					Refer to the "Requirements" section at the top of this file for information
;					on enabling the SAP GUI Scripting interface.
;
; Related .......:
; Link ..........:
; Example .......:	Yes
;
; ;==========================================================================================
Func _ConfluenceDomainSet($domain)

	$confluence_domain = $domain
EndFunc

; #FUNCTION# ;===============================================================================
;
; Name...........:	_InsightsLogin()
; Description ...:	Login a user to Janison Insights.
; Syntax.........:	_InsightsLogin($username, $password)
; Parameters ....:	$win_title			- Optional: The title of the SAP window (within the session) to attach to.
;											The window "SAP Easy Access" is used if one isn't provided.
;											This may be a substring of the full window title.
;					$sap_transaction	- Optional: a SAP transaction to run after attaching to the session.
;											A "/n" will be inserted at the beginning of the transaction
;											if one isn't provided.
; Return values .: 	On Success			- Returns True.
;                 	On Failure			- Returns False, and:
;											sets @ERROR = 1 if unable to find an active SAP session.
;												This means the SAP GUI Scripting interface is not enabled.
;												Refer to the "Requirements" section at the top of this file.
;											sets @ERROR = 2 if unable to find the SAP window to attach to.
;
; Author ........:	seangriffin
; Modified.......:
; Remarks .......:	A prerequisite is that the SAP GUI Scripting interface is enabled,
;					and the SAP user is already logged in (ie. The "SAP Easy Access" window is displayed).
;					Refer to the "Requirements" section at the top of this file for information
;					on enabling the SAP GUI Scripting interface.
;
; Related .......:
; Link ..........:
; Example .......:	Yes
;
; ;==========================================================================================
Func _ConfluenceLogin($username, $password)

	$confluence_username = $username
	$confluence_password = $password
EndFunc


; Content

Func _ConfluenceGetPageVersion($page_key)

	Local $iPID = Run('curl.exe -k -H "Accept: application/json" -H "Content-Type: application/json" -X GET -u ' & $confluence_username & ':' & $confluence_password & ' ' & $confluence_domain & '/wiki/rest/api/content/' & $page_key & '?expand=version', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $confluence_json = StdoutRead($iPID)

	$rr = StringRegExp($confluence_json, '(?U)"number":(.*),', 3)
	Return Int($rr[0])

EndFunc

Func _ConfluenceGetNextPageVersion($page_key)

	Local $current_version = _ConfluenceGetPageVersion($page_key)
	Return $current_version + 1

EndFunc

Func _ConfluenceCreatePage($space_key, $ancestor_key, $title, $body)

	Local $results_json = '{"type":"page","title":"' & $title & '","space":{"key":"' & $space_key & '"},"ancestors":[{"id":"' & $ancestor_key & '"}],"body":{"storage":{"value":"' & $body & '","representation":"storage"}}}'
	FileDelete(@ScriptDir & "\curl_in.json")
	FileWrite(@ScriptDir & "\curl_in.json", $results_json)

	Local $iPID = Run('curl.exe -k -H "Accept: application/json" -H "Content-Type: application/json" -X POST --data @curl_in.json -u ' & $confluence_username & ':' & $confluence_password & ' ' & $confluence_domain & '/wiki/rest/api/content', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $confluence_json = StdoutRead($iPID)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $confluence_json = ' & $confluence_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
EndFunc

Func _ConfluenceUpdatePage($space_key, $ancestor_key, $page_key, $title, $body)

	Local $next_version = _ConfluenceGetNextPageVersion($page_key)

	Local $results_json = '{"version":{"number":' & $next_version & '},"type":"page","title":"' & $title & '","space":{"key":"' & $space_key & '"},"ancestors":[{"id":"' & $ancestor_key & '"}],"body":{"storage":{"value":"' & $body & '","representation":"storage"}}}'
	FileDelete(@ScriptDir & "\curl_in.json")
	FileWrite(@ScriptDir & "\curl_in.json", $results_json)

	Local $iPID = Run('curl.exe -k -H "Accept: application/json" -H "Content-Type: application/json" -X PUT --data @curl_in.json -u ' & $confluence_username & ':' & $confluence_password & ' ' & $confluence_domain & '/wiki/rest/api/content/' & $page_key, @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    $confluence_json = StdoutRead($iPID)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $confluence_json = ' & $confluence_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
EndFunc

Func _ConfluenceWikiToHtml($wiki_markup)

	FileDelete(@ScriptDir & "\curl_in.json")
	FileWrite(@ScriptDir & "\curl_in.json", $wiki_markup)

	$curl = 'curl.exe -k -H "Accept: application/json" -H "Content-Type: application/json" -X POST --data @curl_in.json -u ' & $confluence_username & ':' & $confluence_password & ' ' & $confluence_domain & '/rest/tinymce/1/wikixhtmlconverter'
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $curl = ' & $curl & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

	Local $iPID = Run('curl.exe -k -H "Accept: application/json" -H "Content-Type: application/json" -X POST --data @curl_in.json -u ' & $confluence_username & ':' & $confluence_password & ' ' & $confluence_domain & '/rest/tinymce/1/wikixhtmlconverter', @ScriptDir, @SW_HIDE, $STDOUT_CHILD)
    ProcessWaitClose($iPID)
    local $confluence_html = StdoutRead($iPID)
	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $confluence_html = ' & $confluence_html & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
	return $confluence_html
EndFunc




Func SQLite_to_Confluence_Chart($type, $title, $subtitle, $xlabel, $ylabel, $legend, $stacked, $width, $height, $show_shapes, $opacity, $data_display, $tables, $columns, $data_orientation, $time_series, $date_format, $time_period, $language, $country, $forgive, $colors, $range_axis_lower_bound, $range_axis_upper_bound, $category_label_position, $query)

	$storage_format = $storage_format & '<ac:structured-macro ac:name=\"chart\">' & @CRLF

	if StringLen($type) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"type\">' & $type & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($title) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"title\">' & $title & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($subtitle) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"subTitle\">' & $subtitle & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($xlabel) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"xLabel\">' & $xlabel & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($ylabel) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"yLabel\">' & $ylabel & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($legend) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"legend\">' & $legend & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($stacked) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"stacked\">' & $stacked & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($width) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"width\">' & $width & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($height) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"height\">' & $height & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($show_shapes) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"showShapes\">' & $show_shapes & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($opacity) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"opacity\">' & $opacity & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($data_display) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"dataDisplay\">' & $data_display & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($tables) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"tables\">' & $tables & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($columns) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"columns\">' & $columns & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($data_orientation) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"dataOrientation\">' & $data_orientation & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($time_series) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"timeSeries\">' & $time_series & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($date_format) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"dateFormat\">' & $date_format & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($time_period) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"timePeriod\">' & $time_period & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($language) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"language\">' & $language & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($country) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"country\">' & $country & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($forgive) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"forgive\">' & $forgive & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($colors) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"colors\">' & $colors & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($range_axis_lower_bound) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"rangeAxisLowerBound\">' & $range_axis_lower_bound & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($range_axis_upper_bound) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"rangeAxisUpperBound\">' & $range_axis_upper_bound & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($category_label_position) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"categoryLabelPosition\">' & $category_label_position & '</ac:parameter>' & @CRLF
	EndIf

	Local $double_quotes = """"

	if $confluence_html = true Then

		$double_quotes = "\"""
	EndIf

	$iRval = _SQLite_GetTable2d(-1, $query, $aResult, $iRows, $iColumns)

	If $iRval = $SQLITE_OK Then

;		_SQLite_Display2DResult($aResult)

		Local $num_rows = UBound($aResult, 1)
		Local $num_cols = UBound($aResult, 2)

		if $num_rows < 2 Then

	;		$storage_format = $storage_format &	"<p>" & $empty_message & "</p>" & @CRLF
		Else

			$storage_format = $storage_format &	"<ac:rich-text-body>" & @CRLF
			$storage_format = $storage_format &	"<table>" & @CRLF
			$storage_format = $storage_format &	"<tbody>" & @CRLF
			$storage_format = $storage_format & "<tr>" & @CRLF

			for $i = 0 to ($num_cols - 1)

				$storage_format = $storage_format & "<th>" & $aResult[0][$i] & "</th>" & @CRLF
			Next

			$storage_format = $storage_format & "</tr>" & @CRLF

			for $i = 1 to ($num_rows - 1)

				$storage_format = $storage_format & "<tr>"

				for $j = 0 to ($num_cols - 1)

;						$aResult[$i][$j] = StringReplace($aResult[$i][$j], " \</td>", " \\</td>")
					$aResult[$i][$j] = StringRegExpReplace($aResult[$i][$j], "([^\\])\\$", "$1\\\\")
;						$a = StringRegExpReplace($a, "([^\\])\\$", "$1\\\\")
					$aResult[$i][$j] = StringReplace($aResult[$i][$j], "<br>", "<br/>")
					$aResult[$i][$j] = StringReplace($aResult[$i][$j], "&", "&amp;")
					$aResult[$i][$j] = StringReplace($aResult[$i][$j], """", "\""")
					$aResult[$i][$j] = StringReplace($aResult[$i][$j], "\\""", "\""")
					$storage_format = $storage_format & "<td>" & $aResult[$i][$j] & "</td>" & @CRLF
				Next

				$storage_format = $storage_format & "</tr>" & @CRLF
			Next

			$storage_format = $storage_format &	"</tbody>" & @CRLF
			$storage_format = $storage_format &	"</table>" & @CRLF
			$storage_format = $storage_format &	"</ac:rich-text-body>" & @CRLF
		EndIf
	Else
		MsgBox($MB_SYSTEMMODAL, "SQLite Error: " & $iRval, _SQLite_ErrMsg())
	EndIf

	$storage_format = $storage_format &	"</ac:structured-macro><br /><br />" & @CRLF

EndFunc




Func SQLite_with_weeks_to_Confluence_Chart($type, $title, $subtitle, $xlabel, $ylabel, $orientation, $legend, $stacked, $width, $height, $show_shapes, $opacity, $data_display, $tables, $columns, $data_orientation, $time_series, $date_format, $time_period, $language, $country, $forgive, $colors, $range_axis_lower_bound, $range_axis_upper_bound, $range_axis_tick_unit, $category_label_position, $query)

	$storage_format = $storage_format & '<ac:structured-macro ac:name=\"chart\">' & @CRLF

	if StringLen($type) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"type\">' & $type & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($title) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"title\">' & $title & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($subtitle) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"subTitle\">' & $subtitle & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($xlabel) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"xLabel\">' & $xlabel & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($ylabel) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"yLabel\">' & $ylabel & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($orientation) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"orientation\">' & $orientation & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($legend) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"legend\">' & $legend & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($stacked) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"stacked\">' & $stacked & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($width) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"width\">' & $width & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($height) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"height\">' & $height & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($show_shapes) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"showShapes\">' & $show_shapes & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($opacity) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"opacity\">' & $opacity & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($data_display) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"dataDisplay\">' & $data_display & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($tables) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"tables\">' & $tables & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($columns) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"columns\">' & $columns & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($data_orientation) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"dataOrientation\">' & $data_orientation & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($time_series) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"timeSeries\">' & $time_series & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($date_format) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"dateFormat\">' & $date_format & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($time_period) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"timePeriod\">' & $time_period & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($language) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"language\">' & $language & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($country) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"country\">' & $country & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($forgive) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"forgive\">' & $forgive & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($colors) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"colors\">' & $colors & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($range_axis_lower_bound) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"rangeAxisLowerBound\">' & $range_axis_lower_bound & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($range_axis_upper_bound) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"rangeAxisUpperBound\">' & $range_axis_upper_bound & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($range_axis_tick_unit) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"rangeAxisTickUnit\">' & $range_axis_tick_unit & '</ac:parameter>' & @CRLF
	EndIf

	if StringLen($category_label_position) > 0 Then

		$storage_format = $storage_format & '<ac:parameter ac:name=\"categoryLabelPosition\">' & $category_label_position & '</ac:parameter>' & @CRLF
	EndIf

	Local $double_quotes = """"

	if $confluence_html = true Then

		$double_quotes = "\"""
	EndIf

	$iRval = _SQLite_GetTable2d(-1, $query, $aResult, $iRows, $iColumns)

	If $iRval = $SQLITE_OK Then

;		_SQLite_Display2DResult($aResult)

		Local $num_rows = UBound($aResult, 1)
		Local $num_cols = UBound($aResult, 2)

		if $num_rows < 2 Then

	;		$storage_format = $storage_format &	"<p>" & $empty_message & "</p>" & @CRLF
		Else

			$storage_format = $storage_format &	"<ac:rich-text-body>" & @CRLF
			$storage_format = $storage_format &	"<table>" & @CRLF
			$storage_format = $storage_format &	"<tbody>" & @CRLF
			$storage_format = $storage_format & "<tr>" & @CRLF

			for $i = 0 to ($num_cols - 1)

				$storage_format = $storage_format & "<th>" & $aResult[0][$i] & "</th>" & @CRLF
			Next

			$storage_format = $storage_format & "</tr>" & @CRLF

			for $i = 1 to ($num_rows - 1)

				$storage_format = $storage_format & "<tr>"

				for $j = 0 to ($num_cols - 1)

					if $j = 0 Then

						$aResult[$i][$j] = _DateFromWeekNumber(2019, ($aResult[$i][$j] + 1))
					Else

	;						$aResult[$i][$j] = StringReplace($aResult[$i][$j], " \</td>", " \\</td>")
						$aResult[$i][$j] = StringRegExpReplace($aResult[$i][$j], "([^\\])\\$", "$1\\\\")
	;						$a = StringRegExpReplace($a, "([^\\])\\$", "$1\\\\")
						$aResult[$i][$j] = StringReplace($aResult[$i][$j], "<br>", "<br/>")
						$aResult[$i][$j] = StringReplace($aResult[$i][$j], "&", "&amp;")
						$aResult[$i][$j] = StringReplace($aResult[$i][$j], """", "\""")
						$aResult[$i][$j] = StringReplace($aResult[$i][$j], "\\""", "\""")
					EndIf

					$storage_format = $storage_format & "<td>" & $aResult[$i][$j] & "</td>" & @CRLF
				Next

				$storage_format = $storage_format & "</tr>" & @CRLF
			Next

			$storage_format = $storage_format &	"</tbody>" & @CRLF
			$storage_format = $storage_format &	"</table>" & @CRLF
			$storage_format = $storage_format &	"</ac:rich-text-body>" & @CRLF
		EndIf
	Else
		MsgBox($MB_SYSTEMMODAL, "SQLite Error: " & $iRval, _SQLite_ErrMsg())
	EndIf

	$storage_format = $storage_format &	"</ac:structured-macro><br /><br />" & @CRLF

EndFunc


Func Update_Confluence_Page($url, $space_key, $ancestor_key, $page_key, $page_title, $page_body)

	_ConfluenceSetup()
	_ConfluenceDomainSet($url)
	_ConfluenceLogin($confluence_username, $confluence_decrypted_password)
	_ConfluenceUpdatePage($space_key, $ancestor_key, $page_key, $page_title, $page_body)
	_ConfluenceShutdown()

EndFunc


; The week with the first Thursday of the year is week number 1.
Func _DateFromWeekNumber($iYear, $iWeekNum)
    Local $Date, $sFirstDate = _DateToDayOfWeek($iYear, 1, 1)
    If $sFirstDate < 6 Then
        $Date = _DateAdd("D", 2 - $sFirstDate, $iYear & "/01/01")
    ElseIf $sFirstDate = 6 Then
        $Date = _DateAdd("D", $sFirstDate - 3, $iYear & "/01/01")
    ElseIf $sFirstDate = 7 Then
        $Date = _DateAdd("D", $sFirstDate - 5, $iYear & "/01/01")
    EndIf
    ;ConsoleWrite(_DateToDayOfWeek($iYear, 1, 1) &"  ")
    Local $aDate = StringSplit($Date, "/", 2)
    Return _DateAdd("w", $iWeekNum - 1, $aDate[0] & "/" & $aDate[1] & "/" & $aDate[2])
EndFunc   ;==>_DateFromWeekNumber


Func _ConfluenceAuthenticationWithToast($app_name, $domain, $ini_filename)

	$confluence_username = IniRead($ini_filename, "main", "confluenceusername", "")
	$confluence_encrypted_password = IniRead($ini_filename, "main", "confluencepassword", "")

	_JiraSetup()
	_JiraDomainSet($domain)
	_Toast_Set(0, -1, -1, -1, -1, -1, "", 100, 100)
	_Toast_Show(0, $app_name, "Login to Confluence ...", -300, False, True)

	if stringlen($confluence_encrypted_password) > 0 Then

		$confluence_decrypted_password = _Crypt_DecryptData($confluence_encrypted_password, @ComputerName & @UserName, $CALG_AES_256)
		$confluence_decrypted_password = BinaryToString($confluence_decrypted_password)
	EndIf

	if stringlen($confluence_decrypted_password) > 0 Then

		_ConfluenceLogin($confluence_username, $confluence_decrypted_password)
		_JiraLogin($confluence_username, $confluence_decrypted_password)
		_JiraGetCurrentUser()
	EndIf

	if stringlen($confluence_decrypted_password) = 0 or StringInStr($jira_json, "<title>Unauthorized (401)</title>", 1) > 0 Then

		_Toast_Show(0, $app_name, "Username or password incorrect or not set.                       " & @CRLF & "Set your Confluence login below." & @CRLF & @CRLF & @CRLF & @CRLF & @CRLF, -9999, False, True)
		GUICtrlCreateLabel("Username:", 10, 70, 80, 20)
		Local $username_input = GUICtrlCreateInput("", 80, 70, 200, 20)
		GUICtrlCreateLabel("Password:", 10, 90, 80, 20)
		Local $password_input = GUICtrlCreateInput("", 80, 90, 200, 20, $ES_PASSWORD)
		$done_button = GUICtrlCreateButton("Done", 80, 110, 80, 20)

		While 1

			$msg = GUIGetMsg()

			if $msg = $done_button Then

				$confluence_username = GUICtrlRead($username_input)
				Local $confluence_decrypted_password = GUICtrlRead($password_input)
				_Toast_Show(0, $app_name, "Login to Confluence ...", -300, False, True)
				_ConfluenceLogin($confluence_username, $confluence_decrypted_password)
				_JiraLogin($confluence_username, $confluence_decrypted_password)
				_JiraGetCurrentUser()

				ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $jira_json = ' & $jira_json & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

				if StringInStr($jira_json, "<title>Unauthorized (401)</title>", 1) = 0 Then

					IniWrite($ini_filename, "main", "confluenceusername", $confluence_username)
					$confluence_encrypted_password = _Crypt_EncryptData($confluence_decrypted_password, @ComputerName & @UserName, $CALG_AES_256)
					IniWrite($ini_filename, "main", "confluencepassword", $confluence_encrypted_password)
				EndIf

				_Toast_Hide()
				ExitLoop
			EndIf

			if $hToast_Handle = 0 Then

				Exit
			EndIf
		WEnd
	EndIf

	if StringInStr($jira_json, "<title>Unauthorized (401)</title>", 1) > 0 Then

		_Toast_Show(0, $app_name, "Username or password incorrect or not set." & @CRLF & "Exiting ...", -5, true, True)
		Exit
	EndIf

EndFunc
