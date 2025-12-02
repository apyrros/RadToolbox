; =====================================================================================
; RadToolbox.ahk  (AutoHotkey v1)
; -------------------------------------------------------------------------------------
; Purpose:
;   Compact, portable AI assistant script for radiology reporting demonstrations.
;   Works with Nuance PowerScribe 360 and Epic Hyperspace style text boxes by sending
;   keystrokes and clipboard operations (no vendor SDK required).
;
; Supported providers:
;   - Azure OpenAI (chat/completions endpoint)
;   - OpenAI API (chat/completions)
;   - Ollama local models (http://localhost:11434/api/chat)
;
; Features retained (only AI tools – all site-specific items removed):
;   - Generate Impression from the active PowerScribe report.
;   - Check Report for Errors with a clickable error terminal.
;   - Restore Dictation via a local history buffer.
;   - Pull Indication from Epic / clipboard and clean it with GPT.
;   - Prompt Manager with editable built-in prompts and Custom Prompt menu.
;   - Apply Prompt to Selection (highlight → transform in-place).
;   - AI-enabled Notes (Epic copy/paste helper) with optional enable/disable.
;   - Preferences dialog for provider, endpoint, API key, hotkeys, toggles.
;   - Startup disclaimer (EULA-style) shown once per session.
;
; Safe defaults and portability:
;   - No hard-coded API keys; keys are entered by the user and stored per-user under
;     %A_AppData%\AI_Tools_Demo\settings.ini.
;   - No fixed drive letters; all paths use AutoHotkey built-ins (A_AppData, A_Temp,
;     A_ScriptDir).
;   - Single self-contained file; no external includes required.
;
; IMPORTANT NOTES FOR DEMO USE (edit as needed for your environment):
;   - PowerScribe integration assumes the active window is "PowerScribe 360". If your
;     window title differs, change PowerScribeWindowTitle below.
;   - Epic integration looks for "Hyperspace" in the title by default.
;   - Default model names and prompts can be edited in the Preferences dialog or by
;     adjusting the DefaultSettings / DefaultPrompts sections.
;   - API calls run over HTTPS/HTTP; if a call fails, a clear message box is shown and
;     the original dictation remains available in the history buffer.
; =====================================================================================

#NoEnv
; --- Embedded JSON.ahk (cJson 0.4.1) ---
;
; cJson.ahk 0.4.1
; Copyright (c) 2021 Philip Taylor (known also as GeekDude, G33kDude)
; https://github.com/G33kDude/cJson.ahk
;
; MIT License
;
; Permission is hereby granted, free of charge, to any person obtaining a copy
; of this software and associated documentation files (the "Software"), to deal
; in the Software without restriction, including without limitation the rights
; to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
; copies of the Software, and to permit persons to whom the Software is
; furnished to do so, subject to the following conditions:
;
; The above copyright notice and this permission notice shall be included in all
; copies or substantial portions of the Software.
;
; THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
; IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
; FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
; AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
; LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
; OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
; SOFTWARE.
;

class JSON
{
	static version := "0.4.1-git-built"

	BoolsAsInts[]
	{
		get
		{
			this._init()
			return NumGet(this.lib.bBoolsAsInts, "Int")
		}

		set
		{
			this._init()
			NumPut(value, this.lib.bBoolsAsInts, "Int")
			return value
		}
	}

	EscapeUnicode[]
	{
		get
		{
			this._init()
			return NumGet(this.lib.bEscapeUnicode, "Int")
		}

		set
		{
			this._init()
			NumPut(value, this.lib.bEscapeUnicode, "Int")
			return value
		}
	}

	_init()
	{
		if (this.lib)
			return
		this.lib := this._LoadLib()

		; Populate globals
		NumPut(&this.True, this.lib.objTrue, "UPtr")
		NumPut(&this.False, this.lib.objFalse, "UPtr")
		NumPut(&this.Null, this.lib.objNull, "UPtr")

		this.fnGetObj := Func("Object")
		NumPut(&this.fnGetObj, this.lib.fnGetObj, "UPtr")

		this.fnCastString := Func("Format").Bind("{}")
		NumPut(&this.fnCastString, this.lib.fnCastString, "UPtr")
	}

	_LoadLib32Bit() {
		static CodeBase64 := ""
		. "FLYQAQAAAAEwVYnlEFOB7LQAkItFFACIhXT///+LRUAIixCh4BYASAAgOcIPhKQAcMdFAvQAFADrOIN9DAAAdCGLRfQF6AEAQA+2GItFDIsAAI1I"
		. "AotVDIkACmYPvtNmiRAg6w2LRRAAKlABwQAOiRCDRfQAEAViIACEwHW5AMaZiSBFoIlVpAEhRCQmCABGAAYEjQATBCSg6CYcAAACaRQLXlDHACIA"
		. "DFy4AZfpgK0HAADGRfMAxAgIi1AAkwiLQBAQOcJ1RwATAcdFCuwCuykCHAyLRewAweAEAdCJRbACiwABQAiLVeyDAMIBOdAPlMCIAEXzg0XsAYB9"
		. "EPMAdAuEIkXsfIrGgkUkAgsHu1sBJpgFu3uCmYlOiRiMTQSAvYGnAHRQx0Wi6Auf6AX5KJ/oAAQjhRgCn8dF5AJ7qQULgUGDauSEaqyDfeSwAA+O"
		. "qYAPE6EsDaGhhSlSx0XgiyngqilO4AACRQyCKesnUyAgIVUgZcdF3EIgVMdERdiLItgF/Kgi2EcAAkUMgiKDRdyABBiAO0XcfaQPtoB5gPABhMAP"
		. "hJ/AwIHCeRg5ReR9fOScGItFrMCNALCYiVVKnA2wmAGwZRlEXxfNDxPpgTjKE+nKQgSAIaIcgCEPjZ9C3NQLQOjUBf4oQNQAAkUMRNyhxCyQiVWU"
		. "zSyQYRaUsRiYbivqC+scwwlgi1UQiVTgCOAEVKQkBIEIYBqVCDqtKAN/Q4ctDIP4AXUek0EBLg76FwBhnAIBKKEDBQYPhV7COqzAmIbkICAAgVXH"
		. "RdDLKbjQBQcAB98pwynQAAETJQbCKekqJA4QodzFRgzMSwzMBQxfDEYM7swAASUGQwzHphiBsUMMYshLDMgFEl8MRgzIFwABJQZDDGRCDBiNSBAB"
		. "D7aVg7+siwAoiUwkoSwMjy3N+TD//+kv5BKBLQV1liBCBk8FVsBJ6QRIBYgCdWlAAY1VgCUEVNQUwVzEIho3IhogAItViItFxAHAiI0cAioaD7cT"
		. "ERoExAEFBgHQD7cAgGaFwHW36ZCiZ2LACyXABRcfJeYKwM8AASUGpmcuHIAVv9RGCgbkAAHjyeQPjEj6RP//ZJ4PhLXiFbzt6xW8P6+IC7wAASUG"
		. "BMRv4uLhqGH7CAu0/6iIBbQXgAAVA3RUuCABuDtFqBh8pFpxXVNxfV9xA11xkgmLXfzJw5DOkLEaAgBwiFdWkIizUYoMMBQUcQDHQAjRAdjHQAxS"
		. "DIAECIEEwCEOCJFBwABhH4P4IHQi5dgACnTX2AANdCLJ2AAJdLvYAHsPjIVygjxoBcdFoDIHVkWBj2AAqGMArGEAoYaM8AjQLkAYixWhAJDHRCQg"
		. "4gFEJCCLIAAAjU2QwDMYjdRNoGAAFFABEEGWcAAf8gtwAOMMQFdxAIkUJED/0IPsJItAY0U+sN8N3w3fDd8N1wB9D6yEVARuEgGFEG9DCQFAg/gi"
		. "dAq4YCj/xOm/EAqNRYDxYOEHAeAtaf7//4XAdPoX8wGf8AH/Cf8J/wn/CXXVADrFB0LPBZJplAjfVv2SCMQCFcICiIM4CP1jArCyZ4ABTxRPCk8K"
		. "TwqR1wAsdRIqBelUcBFmkFkWhQl8C18MgCwJQQIxVbCJUAjDqlTDdQLzA1sPhfBFGTYovIVwwUGxIjK5kwB4lgDOfJQA/yj8KI1gkAIiKZ6NEQVf"
		. "KV8pVimFaBED/EW00KbxAq8VrxWvFa8VYdcAXQ+EtpSP9imlkwNA2B/h+9kfFwr1AdXgi+RjArRhArpQFS8Kby8KLwovCtkfFioFgVzplgGACBkg"
		. "XcUJegkfIDUXILQWIFJ1AkQ4D4VMYwPvNYB4ReCSA+DDkAOjBAgA6e8FSxRvDbQH/pEgNwVcD4WqF51NKQdxe+CAAYlV4LsCazsuizkGwATbAlzc"
		. "Aqpd2wIv2wIv3AIv2wKqYtsCCNwCAdsCZtsCqgzcAtPbTW7bAgrcAqql2wJy2wIN3AJ32wIudNsCMR7ZAknbAnUPfIURTT7gA4ADsWVCz+nPwdcw"
		. "AQADoNyJwuEBOhuIL34w2AA5fyLDAoORAlMBAdCD6DCFAwTpgKk1g/hAfi0B2ADAtwBGfx+LReAPtwAAicKLRQiLAEEAkAHQg+g3AXDgIGaJEOtF"
		. "BVhmg1D4YH4tCDRmE+hXEQZ0Crj/AADpbQZEAAACQI1QAgAOiQAQg0XcAYN93BADD44WAD6DReAoAusmAypCBCoQjQpKAioIAEmNSAKJGk0AZhIA"
		. "Ugh9Ig+FAP/8//+LRQyLEkgBJinIAXcMi0AQCIPoBAEp4GbHCgAMeLgAEADp3QUjBBYDSC10JIgGLw8IjrEDig85D4+foYAIx0XYAYInDIArIhSB"
		. "A8dACIEnx0DmDAEDiSh1FIAWAWiKPjGIEDB1IxMghRXpjhELKTB+dQlJf2frCkcBdlCBd2vaCmtAyAAB2bsKgBn3AOMB0YnKi00IAIsJjXECi10I"
		. "AIkzD7cxD7/OAInLwfsfAcgRANqDwNCD0v+LAE0MiUEIiVEMSck+fhoJGX6dRXCrEAQAAJCIBi4PhYalTSyGI2YPbsDAAADKZg9iwWYP1mSFUEAQ"
		. "361BAYAI3VZYwGpBUAUAVNQBVOsAQotV1InQweAAAgHQAcCJRdQBQxVIAotVCIkKAcAbmIPoMImFTIXAD9tDAUXU3vmBErBACN7BhRTIMA7KMCKi"
		. "SANldBJIA0UPHIVVACANMQMHFHUxVQk00MAA2gA00wA0lVEVNMZF00uBE0AEAY3KF+tAzAYIK3URhgxX0IhNMsRiH8KizEGM61Ani1XMh07DUU4B"
		. "ENiJRcxYFb3HRSLIwTDHRcRCChOLhFXIqDHIg0XEQBgAxDtFzHzlgH0Q0wB0E0Mv20XIoaMwWAjrEUcCyUYiFeUoKyR0WCBN2JmJAN8Pr/iJ1g+v"
		. "APEB/vfhjQwWk2FVJFHrHcYGBXVmCibYcApELgMAA3oMAqFqZXQPhasiGsAiGgA3i0XABQcXAAAAD7YAZg++0FEmBTnCdGQqy+1AgwxFwKAexgaE"
		. "wHW6lA+2wIYAQAF0G6UPJ0N4oidDeOssQwMJABCLFeQWgoWJUAhCoUIBAItABKMCiYAUJP/Qg+wEgxcuT2UPhKqFF7yFF7wF6gyaFw6PF7yAF8YG"
		. "mhf76I+JF9yHF0IBgxdBAYsXgpKrlG51f8dFIgOA6zSLRbgFEhMX0gcCF+tYrBa4oBZmBvWgFr3nEeDnEUIB4xFBAQnqEesFIguNZfRbMF5fXcNB"
		. "AgUAIlUAbmtub3duX08AYmplY3RfAA0KCiALIqUBdHJ1ZQAAZmFsc2UAbgh1bGzHBVZhbHUAZV8AMDEyMzQANTY3ODlBQkMAREVGAFWJ5VNAg+xU"
		. "x0X0ZreLAEAUjVX0iVQkIBTHRCQQIitEJKIMwUONVQzAAgjAAQ8AqaAF4HPDFhjHReSpAgVF6MMA7MMA8IMKcBCJReRgAuPOIgwYqItV9MAIIKQL"
		. "HOQAghjhAI1N5IlMgw/fwQyBD8QDwjwgEAQnD2De0hCDNgl1MCEQcE7xBUAIi1UQi1JFAgTE62hmAgN1XGECElESu0AWf7lBBTnDGSjRfBWGAT0g"
		. "AYCJQNCD2P99LvAajTRV4HEPiXAPMR4EJATooQAChcB0EYsETeBGA4kBiVEEAJCLXfzJw5CQAXAVg+xYZsdF7ikTH0XwIBYUARBNDAC6zczMzInI"
		. "9xDiweoDNkopwYkCyhAHwDCDbfQBgSGA9GaJVEXGsAMJ4gL34pAC6AOJRQAMg30MAHW5jUJVoAH0AcABkAIQDYAJCGIRwwko/v//hpBACLMdYMdF"
		. "+EIuBhrkRcAKRfjB4AQgAdCJRdgBAUAYwDlF+A+NRPAZAAsKzlEC2PEMRfTGRQDzAIN99AB5B2GQAAH3XfRQHEMM9KC6Z2ZmZkAM6nAJhPgCUnkp"
		. "2InC/wyog23s8gzs8QymngNAwfkficopoAj0AYEGdaWAffMAdAYOQQMhA8dERaYtHXAnpsAAwA5gAtDGRYbrkCXiJotF5I0hjCDQAdAPtzBn5I3S"
		. "DMEWAcgDOnWQOQgCQABmhcB1GSUBDGUmAQYQBQHrEKG8AnQDUIS8AnQHg0XkAQDrh5CAfesAD2aEoWbhH1XYMJnRLemSyiQuQBwhFYyj4gChwxTU"
		. "xkXjgAvcgwvq3IIF1IQL3I8LCAKFC/sjAYoL44ILvAKBC7wCgQtC3IML4wB0D0oL65AYg0X48n1AEFIL8Nf9//9ySLosvz1iABNyQ2Aj6AWBD90A"
		. "3RpdkC7YswGyDsdF4ONjACIbjUXoUCcwAZEH7KGIED3jQBWhAB1BIXXATCQYjU3YBUFCav8MQeVIFUEhCz8LPwvAATES0QAxBIsAADqJIEmfC3+f"
		. "C58LnwufC58Lnws2O2Q9wAnmkgrSNjQKV0l8GIM1AStMfW6NRahoSib2kEBUD+s3gUN0IACLVbCLRfABwFSNHPBqDHSWDHGWE12xzg2hIcBs8CAQ"
		. "wWzw1gEFA2YntzNzPvR60xMA7IN97AB5bYtMTeyPQY9BuDAQBCk60KpOvr4DpkHCBXWjB+ECwQJAQb4tAOtbX88GzwZfVa8GrwalhCPrQj5CEyeN"
		. "Vb7WVuhnvxO/E7IT6AF8AyYUqWvpNbMqGJIGF3oFUIMimADpyXLcmAXpt5PdKQTkdVaiAxStA1wAVx0JHwYTBmMeBlEZBlxPHwYfBh8GaALpAR4G"
		. "73vTaBMGCB8GHwYfBmYCYgAAgrEA6Z8CAACLRRAgiwCNUAEAcIkQBOmNAogID7cAZgCD+Ax1VoN9DEAAdBSLRQwAjEgAAotVDIkKZsdgAFwA6w0K"
		. "3AJMF6ENTGYA6T0OwisJwoIKPGFuAOnbAQ1hFskCEQRhDTxhcgDpKnmOMGeJMAm8MHQAFOkXjjAFgAgPtgUABAAAAITAdCkRBjYfdgyGBX52B0K4"
		. "ABMA6wW4gAIAoIPgAeszCBQYCBTCE4QFPaAAdw0awBc2bykwjgl1jQkDGw+3AMCLVRCJVCQIAQEKVCQEiQQk6DptgR4rwhHAJ8gRi1UhwAwSZokQ"
		. "jRxFCAICBC+FwA+FOvwU//9TISJNIZDJwwCQkJBVieVTgwTsJIAQZolF2McARfAnFwAAx0UC+AE/6y0Pt0XYAIPgD4nCi0XwAAHQD7YAZg++ANCL"
		. "RfhmiVRFQugBB2bB6AQBDoMARfgBg334A36gzcdF9APBDjOCIQAci0X0D7dcRZLoiiOJ2hAybfRAEBD0AHnHAl6LXfwBwic="
		static Code := false
		if ((A_PtrSize * 8) != 32) {
			Throw Exception("_LoadLib32Bit does not support " (A_PtrSize * 8) " bit AHK, please run using 32 bit AHK")
		}
		; MCL standalone loader https://github.com/G33kDude/MCLib.ahk
		; Copyright (c) 2021 G33kDude, CloakerSmoker (CC-BY-4.0)
		; https://creativecommons.org/licenses/by/4.0/
		if (!Code) {
			CompressedSize := VarSetCapacity(DecompressionBuffer, 3935, 0)
			if !DllCall("Crypt32\CryptStringToBinary", "Str", CodeBase64, "UInt", 0, "UInt", 1, "Ptr", &DecompressionBuffer, "UInt*", CompressedSize, "Ptr", 0, "Ptr", 0, "UInt")
				throw Exception("Failed to convert MCLib b64 to binary")
			if !(pCode := DllCall("GlobalAlloc", "UInt", 0, "Ptr", 9092, "Ptr"))
				throw Exception("Failed to reserve MCLib memory")
			DecompressedSize := 0
			if (DllCall("ntdll\RtlDecompressBuffer", "UShort", 0x102, "Ptr", pCode, "UInt", 9092, "Ptr", &DecompressionBuffer, "UInt", CompressedSize, "UInt*", DecompressedSize, "UInt"))
				throw Exception("Error calling RtlDecompressBuffer",, Format("0x{:08x}", r))
			for k, Offset in [33, 66, 116, 385, 435, 552, 602, 691, 741, 948, 998, 1256, 1283, 1333, 1355, 1382, 1432, 1454, 1481, 1531, 1778, 1828, 1954, 2004, 2043, 2093, 2360, 2371, 3016, 3027, 5351, 5406, 5420, 5465, 5476, 5487, 5540, 5595, 5609, 5654, 5665, 5676, 5725, 5777, 5798, 5809, 5820, 7094, 7105, 7280, 7291, 8610, 8949] {
				Old := NumGet(pCode + 0, Offset, "Ptr")
				NumPut(Old + pCode, pCode + 0, Offset, "Ptr")
			}
			OldProtect := 0
			if !DllCall("VirtualProtect", "Ptr", pCode, "Ptr", 9092, "UInt", 0x40, "UInt*", OldProtect, "UInt")
				Throw Exception("Failed to mark MCLib memory as executable")
			Exports := {}
			for ExportName, ExportOffset in {"bBoolsAsInts": 0, "bEscapeUnicode": 4, "dumps": 8, "fnCastString": 2184, "fnGetObj": 2188, "loads": 2192, "objFalse": 5852, "objNull": 5856, "objTrue": 5860} {
				Exports[ExportName] := pCode + ExportOffset
			}
			Code := Exports
		}
		return Code
	}
	_LoadLib64Bit() {
		static CodeBase64 := ""
		. "xrUMAQALAA3wVUiJ5RBIgezAAChIiU0AEEiJVRhMiUUAIESJyIhFKEggi0UQSIsABAWVAh0APosASDnCD0SEvABWx0X8AXrrAEdIg30YAHQtAItF"
		. "/EiYSI0VQo0ATkQPtgQAZkUCGAFgjUgCSItVABhIiQpmQQ++QNBmiRDrDwAbICCLAI1QAQEIiRDQg0X8AQU/TQA/AT4QhMB1pQJ9iUWgEEiLTSAC"
		. "Q41FoABJichIicHoRhYjAI4CeRkQaMcAIgoADmW4gVfpFgkAMADGRfuAZYFsUDBJgwNAIABsdVsADAEox0X0Amw1hBAYiwRF9IBMweAFSAGa0IBG"
		. "sIALgAFQEIALGIPAAQANAImUwIgARfuDRfQBgH2Q+wB0EwEZY9AILRR8sgNWLIIPCEG4wlsBMQZBuHsBuw9gBESJj1+AfSgAdFBkx0XwjLvwgpsm"
		. "mhyxu/DAXcMP5hvHXcjHRezCSqUGAidEQQLsSUGog33sAA8sjsqBL5hhLJQxZsfUReiMMeiCIV+AIa8xluiAMcMPH4gx6y+ZJkIglCZ5x0XkgiZo"
		. "KMdF4Mwo4MIYvhpt8SjgwCjDD37AD8UogwRF5MAFMDtF5H0IkA+2wJDwAYTAuA+E6EDpQVwGkTBBmdyNiZxbUL1AAajgB+FoSpjoaJjkaP4fJQoc"
		. "mTQK6f5DVIkK6epgAo3qEzjiE0Fsx0XcrCay3KIeihm/Jq8m3KAjXeMHSuAHCIOFGpCIGpAthBophhrWJCwsDesbp2YK5AlkCb0gewk6UC4Dv04t"
		. "NItAGIP4AU51YTCAEAoQXB4gcReGA2Iw4wQGD4Wf4EMzYwVhswkYoAHgl2nH1EXYbC/YYicXAAR/L01tL9hgL+MH1xdnL+m0iwJpD21AA2QP1GwP"
		. "utRiB6AABH8PbQ/UYA9V4wdgaQ8Pag8BZw/QtWwP0GIHKn8PcA/QYA8p4wfqFmgPk2JyMI00SAFACk1B48AQAExDgAZBColMJCDBNa1g+P//6WjE"
		. "M8I1Bax1H2QFLDtiITs9SQUQAg+Fg6NtqEiNoJVw////4QSKYJoox0XMIhxIIxwuSIiLlXjAA4tFzAAVYAHATI0EABttHEHoD7cQUxzMkAAKBFBd"
		. "AA+3AGaFwHWeVOmqUjzIHBXIEhHdbhUfFR8V7QbIEBXzA5338ANbPCoRb6AO7zMPTtoFDuzQBahI8XYPjET5RP//8VwPhN3iDMTl7AzE4gjwFO8M"
		. "7wwNB/bEAAfzA7DwA1dzMZRyY+sBkskGvMIChs8GzwbOBha8wAbzA0bIBoNFwIFwAcA7RTB8kKyFOl2khX2vha+FqJFIgeLEAQxdw5AKAOyiDgAK"
		. "VcCjMEEsjawkgBVCpI2zpJURJEiLhQthAKAbFLUASMdACNvyEZAJhaICAQpQAArTAAcRUXUBMSmD+CB01REtAQp0wi0BDXSvES0BCXScLQF7D4WO"
		. "KcJUrweiB8dFUMIQKMdFWHQAYHIAiwWOA+E4AT9BowX1/tAAEMdEJEBTAkQkOAGCAI1VMEiJVCSqMIAAUIEAKJABICG3VEG58QFBkha6ogKJUMFB"
		. "/9LwFzhQbGh/zxDPEM8QzxDPEM8QJwF9WA+EwvJHaQGF8IesgV4Bg/gidAq4IBDw/+lmEYEOoblgB8IeAOj3/f//hcB0+iIDAkUBAu8M7wzvDO8M"
		. "l+8M7wwkAToVCsQQDwi3CAhSKMcLOsMLtAOIsgNJsDKLjQMsRWjESQL/YA1/Go8Njw2PDY8Njw0nAZgsdR1vB2MH6cLQC+dAkIwd1Qy6D58QnBCw"
		. "OQIJtjmLVWhIiVAaCLPSfcoDkwVbD4W+ZUJ4PwX0M/LJcAD4dADTUkIQM8P7+TO10QD/M+yNVdDF8zPw/zP/M+AZwtjwM3DHhay07R8aPx8aHxof"
		. "Gh8aHxonAV0PNoRh45803kdQKCfH+pkpJxUOMQLiJouVcQz1UA1wRCftMBgvDS8NLw0BJAH+tQAKdMJIi4XAAAAAAEiLAA+3AEBmg/gNdK8NkAlE"
		. "dJwNSCx1JAdISAiNUAIFGokQg4UCrAAQAemq/v//gpANbl10Crj/AAC46T4NASoTggAJyAAJMGbHAAkBIwELSIuAVXBIiVAIuAALGADpAQo8A1ki"
		. "D4WMEwUaUwUXiYWgAgkdBFiVggaALQc7CADpRFkEDTGFwHWEXYKCDA8/XA+F9gMhP7mEVnU0AAmCPIETiQJC5YA8IpYg6ccKL4Q6FCOqXBcjgBAj"
		. "L5QRL5cRKjmQEWKUEQiXEfICVY8RZpQRDJcRq5ARblWUEQqXEWSQEXKUEQ11lxEdkBF0lBFCuJMR1sIBjxF1D4WFigWOmcHEFQAAx4WcAcvByw47"
		. "gwyBBoARweAEiUeB/UIKT1MvfkJNAjkcfy/HB2IHxwMB0INk6DDpCemuo2sqCEBEfj9NAkZ/LJoKN6mJCutczQdgLwpmPAqmVyoKhHm1CNcpg0Io"
		. "CAGDvcEAAw+OuIlAmkiDIggC6zrjB8J16QcQSI1K5wchitUjPkggPo0DExJQLmCXLJD7QAtFkkgmBynIBkiCFuMCQAhIg+hOBMs8dRcjpdcHbzEt"
		. "xHQubj4PjgyKp+Q+iA+P9eCgx4WYwSDLh6YADxQGqMdAICCwDDx1IuMGoSTfooMGMHUPITjTCk1+cA4wD46JwdACOX9260yGKAC9AInQSMHgAkgB"
		. "gNBIAcBJicBpDCkgNYuVYwwKoAdID6C/wEwBwGAP0AUISyPFTGYfbg5+jiVMUwgGAAAO4S4PheYD2BtIPmYP78DySIwPKsEUYQLyDxHgQBUGMQXA"
		. "M5TEM+tsixKVYQGJ0MAbAdAB7MCJQgP4G5jAOwIG8AUNcADScAASBGYPKMgQ8g9eyjYHEEAIsPIPWME8CFwQFw8kTI5q6h9jAWV0ngJFuA+F+I9N"
		. "/RCzAhRXImP/Ef8RxoWTDyoBKiFNkwEBTwdDB+syPQMr3HUf3gQfLUsRE68hhCEKOrI1jFRa6zqLlduxAMYbQZ8pnBtEER4xA4NfB18HfqDHhYiE"
		. "IojHhYRVBxyLlVEBSygj4QCDAgIBi2IAOyEyBnzWgL2iD3Qq61kh4BfJUCONUQMQIxoilOsolwJIgxoPKvIFePIPWb0k+R3BpdU6i0FSREiYSA+v"
		. "OTjr8jg6AwV1vwawBqEDvwalugYMtyIDAFNToQ98oPh0D4XfkhOAlROMUouyAJAJjRXSEAOAD7YEEGYPvkEK6ZgDOcIlr0taBZ1moQQL8BYWBYAU"
		. "BYTAdZcAD7YFUuT//4T4wHQdyQqoUtI/FRFkhcwVDgMHV0sF/CI2Q1AIiwXu0QCJwf/SBVMPq/+G+GYPhdMJUQ9FfCIPTItFfN3SCeewAv8O+w5b"
		. "/zz3DmhFfAG1BJu0BJAOoLmQDmjjnw5MYZ4OBKMGbZgO8lQHkw7kggGWDsFBLzP4bg+FpZIOeKESBkmLRXjSCQOfDmWXDgeSDut0bw5lDnhbYA6D"
		. "BLoxJ2MOo+wLVSv4yOMLQ+oLNeoL6wUhUgdIgcQwsAldwz6QBwCkKQ8ADwACACJVAG5rbm93bl9PAGJqZWN0XwANCgoQCSLVAHRydWUAAGZhbHNl"
		. "AG4IdWxs5wJWYWx1AGVfADAxMjM0ADU2Nzg5QUJDAERFRgBVSInlAEiDxIBIiU0QAEiJVRhMiUUgaMdF/ANTRcBREVsoAEiNTRhIjVX8AEiJVCQo"
		. "x0QkEiDxAUG5MSxJicgDcRJgAk0Q/9BIx0RF4NIAx0XodADwwbQEIEiJReDgAFOJAaIFTItQMItF/IpIEAVA0wJEJDiFAOIwggCNVeBGB8BXQAcH"
		. "ogdiFXGWTRBB/9Lz0QWE73UeogaBl8IYYAYT5ADRGOtgpwIDdVODtQEBDIBIOdB9QG4V1AK68Bp/Qhs50H9l4FNF8Q/YSXCIUwfooUE2hcB0D6AB"
		. "2LDuBVADUjAGEJBIg+xmgBge8xXsYPEV5BVmo7IREAWJRfigFhSABACLTRiJyrjNzATMzDBTwkjB6CAgicLB6gMmXinBAInKidCDwDCDzLQAbfwB"
		. "icKLRfwASJhmiVRFwIsARRiJwrjNzMwAzEgPr8JIwegAIMHoA4lFGIMAfRgAdalIjVUDAIQArEgBwEgB0ABIi1UgSYnQSACJwkiLTRDoAQD+//+Q"
		. "SIPEYAhdw5AGAFVIieUASIPscEiJTRAASIlVGEyJRSAQx0X8AAAA6a4CAAAASItFEEiLRFAYA1bB4AUBV4k0RdABD2MAYQEdQDAASDnCD42aAQBg"
		. "AGbHRbgCNAAaQAEAUEXwxkXvAEhAg33wAHkIAAoBAEj3XfDHRegUgwBfAJTwSLpnZgMAgEiJyEj36kgArgDB+AJJichJwWD4P0wpwAG8gQngBgIB"
		. "PABrKcFIicoAidCDwDCDbegVgo3og42QmCdIwflSPwAbSCmBXfACR3WAgIB97wB0EIEigYMhx0RFkC0AgKEGkIIHhKGJRcDGRSDnAMdF4IGJi0Uy"
		. "4IAMjRQBcQEPD7cKEAQJDAEJGEgByAAPtwBmOcJ1b4EPFQBmhcB1HokLi4AXhQsGgDIB6zqTGgR0IlMNdAqDReAQAelm/0B2gH3nkAAPhPYCVkUg"
		. "wH6JwC4QuMBkAOkBQAFlCmw4AWyMysMKhWrIqMZF38A52MM52IYb/sjFOYIE0DmNCsU5xwXLOb7fwjlRDcE5UQ3BOdjGORDfAHQSzTjrIIMsRfwA"
		. "cgg5IAI5O/0M//+ApEA6g8RwXWLDwruB7JABBIS8SGvEdsAB6MQB8MEBwLLgAgUCwPIPEADyD6IRQIXHRcCECMjEAXrQwgGNgGdAioADASNIAIsF"
		. "hOb//0iLoABMi1AwQAN2QQMQx0QkQAMNRCQ4hQICiwAfiVQkMMHtlQECKEAGIAEQQbnBBwpBwi26QgWJwUH/sNJIgcQBF/B3QOl3fwAXABmgeKNs"
		. "gSEACOReD0yJm39veW+4MOAHKRzQgyyTv2+pbw+Feg9gOWEIIwhgb8AtAOkegF8T34IfE9qCx0XsCSEu61DgARgAdDYLi6oAC+xCAUyNBAIzYlRg"
		. "K41IQAFhOQpBQQBlZokQ6w/hU4sQAI1QAQEBiRCDWEXsARQJR2OO5VRAWyc85Dsg6TsDExyvD2aAxwAiAOleBEOAKcgP6UpjAhAhDYP4KCJ1ZmMI"
		. "GXIIXADT7hdcDuYDTw7SYwJEDk5cXw5fDsgF6XNQDl8dSg4IXw5fDsYFYgDp2gBQDuxk5EMODF8OXw5hxgVmAOmNwwsqB3l9KgcKLwcvBy8HLwfi"
		. "Am5IAOkaLwfpBioHDR8vBy8HLwcvB+ICcgDptqcwTy0HkzMBJAcJLwcPLwcvBy8H4gJ0AOk0BS8H6aFXD7YFmdZA//+EwHQr1wcfRHYNxwB+dgcT"
		. "ZwVB4jqD4AHrNqkCGoWpAhTFAD2gAHd9A31ABnxfDV8NXg3vAuECdbPvAtQHD7dRUPFyGCBUUInB6IZxCDTDBB43zwRgAGADEo9MAQhFEANxT0IN"
		. "hcAPhab736BtXwnYQT4EQaggJE71TQtgWdVriQBrjQVC8wdwBVBZxKjrMg+3RXAQg+AP0qzAWlBTtrAAZg++kqiSXugRAjBmwegEEQTRgIN9gPwD"
		. "fsjHRfhwOwgA6z9TCiWLRfjASJhED7dE4HwOC5hEicJfD+BbbfjQBDD4AHm7JVr1Cw=="
		static Code := false
		if ((A_PtrSize * 8) != 64) {
			Throw Exception("_LoadLib64Bit does not support " (A_PtrSize * 8) " bit AHK, please run using 64 bit AHK")
		}
		; MCL standalone loader https://github.com/G33kDude/MCLib.ahk
		; Copyright (c) 2021 G33kDude, CloakerSmoker (CC-BY-4.0)
		; https://creativecommons.org/licenses/by/4.0/
		if (!Code) {
			CompressedSize := VarSetCapacity(DecompressionBuffer, 4249, 0)
			if !DllCall("Crypt32\CryptStringToBinary", "Str", CodeBase64, "UInt", 0, "UInt", 1, "Ptr", &DecompressionBuffer, "UInt*", CompressedSize, "Ptr", 0, "Ptr", 0, "UInt")
				throw Exception("Failed to convert MCLib b64 to binary")
			if !(pCode := DllCall("GlobalAlloc", "UInt", 0, "Ptr", 11168, "Ptr"))
				throw Exception("Failed to reserve MCLib memory")
			DecompressedSize := 0
			if (DllCall("ntdll\RtlDecompressBuffer", "UShort", 0x102, "Ptr", pCode, "UInt", 11168, "Ptr", &DecompressionBuffer, "UInt", CompressedSize, "UInt*", DecompressedSize, "UInt"))
				throw Exception("Error calling RtlDecompressBuffer",, Format("0x{:08x}", r))
			OldProtect := 0
			if !DllCall("VirtualProtect", "Ptr", pCode, "Ptr", 11168, "UInt", 0x40, "UInt*", OldProtect, "UInt")
				Throw Exception("Failed to mark MCLib memory as executable")
			Exports := {}
			for ExportName, ExportOffset in {"bBoolsAsInts": 0, "bEscapeUnicode": 16, "dumps": 32, "fnCastString": 2624, "fnGetObj": 2640, "loads": 2656, "objFalse": 7632, "objNull": 7648, "objTrue": 7664} {
				Exports[ExportName] := pCode + ExportOffset
			}
			Code := Exports
		}
		return Code
	}
	_LoadLib() {
		return A_PtrSize = 4 ? this._LoadLib32Bit() : this._LoadLib64Bit()
	}

	Dump(obj, pretty := 0)
	{
		this._init()
		if (!IsObject(obj))
			throw Exception("Input must be object")
		size := 0
		DllCall(this.lib.dumps, "Ptr", &obj, "Ptr", 0, "Int*", size
		, "Int", !!pretty, "Int", 0, "CDecl Ptr")
		VarSetCapacity(buf, size*2+2, 0)
		DllCall(this.lib.dumps, "Ptr", &obj, "Ptr*", &buf, "Int*", size
		, "Int", !!pretty, "Int", 0, "CDecl Ptr")
		return StrGet(&buf, size, "UTF-16")
	}

	Load(ByRef json)
	{
		this._init()

		_json := " " json ; Prefix with a space to provide room for BSTR prefixes
		VarSetCapacity(pJson, A_PtrSize)
		NumPut(&_json, &pJson, 0, "Ptr")

		VarSetCapacity(pResult, 24)

		if (r := DllCall(this.lib.loads, "Ptr", &pJson, "Ptr", &pResult , "CDecl Int")) || ErrorLevel
		{
			throw Exception("Failed to parse JSON (" r "," ErrorLevel ")", -1
			, Format("Unexpected character at position {}: '{}'"
			, (NumGet(pJson)-&_json)//2, Chr(NumGet(NumGet(pJson), "short"))))
		}

		result := ComObject(0x400C, &pResult)[]
		if (IsObject(result))
			ObjRelease(&result)
		return result
	}

	True[]
	{
		get
		{
			static _ := {"value": true, "name": "true"}
			return _
		}
	}

	False[]
	{
		get
		{
			static _ := {"value": false, "name": "false"}
			return _
		}
	}

	Null[]
	{
		get
		{
			static _ := {"value": "", "name": "null"}
			return _
		}
	}
}


; --- End embedded JSON.ahk ---
#Persistent
#SingleInstance, Force
SendMode, Input
SetWorkingDir, %A_ScriptDir%
SetTitleMatchMode, 2

; -------------------------------------------------------------------------------------
; GLOBAL CONSTANTS / DEFAULTS
; -------------------------------------------------------------------------------------
AppName                    := "AI Tools Demo"
ConfigDir                  := A_AppData "\AI_Tools_Demo"
ConfigFile                 := ConfigDir "\settings.ini"
HistoryLimit               := 10                        ; max undo steps
PowerScribeWindowTitle     := "PowerScribe 360"
EpicWindowHint             := "Hyperspace"
archiveDir                 := ConfigDir "\report_archive"
maxExamplesPerExam         := 20
PromptMessages             := {}
MainbuttonH                := 25
MainbuttonGap              := 6
DefaultEndpoints := { "Azure OpenAI" : "https://<resource>.openai.azure.com/openai/deployments/<deployment>/chat/completions?api-version=2025-01-01-preview"
                    , "OpenAI"       : "https://api.openai.com/v1/chat/completions"
                    , "Ollama"       : "http://localhost:11434/api/chat" }

; Default settings used on first run (modifiable in Preferences)
DefaultSettings := { "Provider"       : "OpenAI"
                   , "ApiKey"         : ""
                   , "Endpoint"       : DefaultEndpoints["OpenAI"]
                   , "Model"          : "gpt-4.1"
                   , "EnableEpicNotes": 1
                   , "AITextColor"    : "0x0080FF"    ; Orange color for AI-generated text (BGR format)
                   , "LastEulaAccept" : "" }          ; session only; persisted timestamp optional

; Built-in prompts (editable in Prompt Manager). These seed the Custom Prompt menu.
DefaultPrompts := [ { "Name": "Academic rewrite", "Text": "Rewrite this text as an academic radiologist. Preserve all clinical details and measurements." }
                  , { "Name": "Fix voice errors", "Text": "Correct voice transcription errors while preserving medical meaning and measurements." }
                  , { "Name": "Summarize for impression", "Text": "Summarize this finding in 1-2 sentences suitable for an impression section." } ]

; Default hotkeys (editable in Preferences)
DefaultHotkeys := { "GenerateImpression" : "^!i"   ; Ctrl+Alt+I
                  , "CheckErrors"         : "^!e"   ; Ctrl+Alt+E
                  , "RestoreDictation"    : "^!r"   ; Ctrl+Alt+R
                  , "CustomPrompt"        : "^!p"   ; Ctrl+Alt+P
                  , "EpicNotes"           : "^!n" } ; Ctrl+Alt+N

; Preset AI text colors (BGR values)
ColorPresets := [ {"Name": "Orange (default)", "BGR": "0x0080FF"}   ; RGB FF8000
                , {"Name": "Bright Blue",      "BGR": "0xFF6600"}   ; RGB 0066FF
                , {"Name": "Green",            "BGR": "0x66CC00"}   ; RGB 00CC66
                , {"Name": "Purple",           "BGR": "0xFF33AA"}   ; RGB AA33FF
                , {"Name": "Red",              "BGR": "0x3333FF"}   ; RGB FF3333
                , {"Name": "Black",            "BGR": "0x000000"} ]

; -------------------------------------------------------------------------------------
; GLOBAL STATE (initialized at runtime)
; -------------------------------------------------------------------------------------
Settings          := {}
Prompts           := []    ; user + built-in prompts
Hotkeys           := {}
HistoryStack      := []    ; undo buffer for dictation
ErrorTerminalData := []    ; stores lines shown in error GUI
SessionEulaShown  := false
GPT_Impression_Numbered := 0  ; numbering toggle for impressions
ErrCtrlMap        := {}    ; control HWNDs that should not trigger drag
PromptHotkeyMap   := {}    ; combo -> prompt index map for custom prompt hotkeys

; -------------------------------------------------------------------------------------
; STARTUP
; -------------------------------------------------------------------------------------
EnsureConfigDir()
LoadSettings()
LoadPrompts()
LoadHotkeys()
BuildMenus()
RegisterAllHotkeys()
ShowStartupEula()
BuildMainGui()
FileCreateDir, %archiveDir%
PromptMessages := BuildPromptMap()  ; Load few-shot examples into cache
return

; =====================================================================================
; MAIN GUI AND MENUS
; =====================================================================================
BuildMainGui(){
    global Settings, AppName, MainMenu, StatusLine
    Gui, Main:New, +AlwaysOnTop +Resize +MinSize400x200, %AppName%
    Gui, Main:Color, 121212, 1c1c1c
    Gui, Main:Margin, 12, 12
    Gui, Main:Font, s10 cFFFFFF, Segoe UI

    Gui, Main:Add, Text, w360, AI tools for PowerScribe / Epic. Use toolbar buttons or the menu bar to trigger GPT actions.

    btnW := 180, btnH := 30, gap := 8
    Gui, Main:Font, s9
    Gui, Main:Add, Button, xm y+12 w%btnW% h%btnH% gGenerateImpression, Generate Impression
    Gui, Main:Add, Button, x+%gap% yp w%btnW% h%btnH% gCheckReportErrors, Check Report for Errors

    Gui, Main:Add, Button, xm y+gap w%btnW% h%btnH% gRestoreDictation, Restore Dictation
    Gui, Main:Add, Button, x+%gap% yp w%btnW% h%btnH% gPullIndication, Pull Indication (Epic)

    Gui, Main:Add, Button, xm y+gap w%btnW% h%btnH% gShowPromptManager, Prompt Manager
    Gui, Main:Add, Button, x+%gap% yp w%btnW% h%btnH% gShowPreferences, Preferences / Settings

    Gui, Main:Add, Button, xm y+gap w%btnW% h%btnH% gRunEpicNote, AI Notes (Epic)
    Gui, Main:Add, Button, x+%gap% yp w%btnW% h%btnH% gOpenFewShotBuilder, Few-Shot Prompt impressionbuilder

	Gui, Main:Font, cFFFFFF
	StatusLine := ""  ; ensure global exists for the GUI control
	Gui, Main:Add, Text, xm y+12 w420 vStatusLine, % "Provider: " Settings.Provider " | Model: " Settings.Model " | Endpoint: " Settings.Endpoint
	Gui, Main:Add, Text, xm y+6 w420, Tips: highlight text then choose a Custom Prompt to transform in-place. Hotkeys are listed under the Hotkeys menu.

    Gui, Main:Menu, MainMenu
    Gui, Main:Show,, %AppName%
}

BuildMenus(){
    global MainMenu, AiMenu, HotkeyMenu

    ; Build AI submenu first (must exist before linking into MainMenu)
    Menu, AiMenu, Add, Generate Impression, GenerateImpression
    Menu, AiMenu, Add, Check Report for Errors, CheckReportErrors
    Menu, AiMenu, Add, Restore Dictation, RestoreDictation
    Menu, AiMenu, Add, Pull Indication (Epic), PullIndication
    Menu, AiMenu, Add
    Menu, AiMenu, Add, Prompt Manager, ShowPromptManager
    Menu, AiMenu, Add, Few-Shot Prompt impressionbuilder, OpenFewShotBuilder
    Menu, AiMenu, Add, Preferences / Settings, ShowPreferences
    Menu, AiMenu, Add
    RebuildCustomPromptMenu()
    Menu, AiMenu, Add, Custom Prompts, :CustomPromptMenu

    ; Hotkey submenu
    Menu, HotkeyMenu, Add, Show Hotkey List, ShowHotkeyHelp
    Menu, HotkeyMenu, Add, Edit Hotkeys (Preferences), ShowPreferences

    ; Main menu bar
    Menu, MainMenu, Add, AI Tools, :AiMenu
    Menu, MainMenu, Add, Hotkeys, :HotkeyMenu
    Menu, MainMenu, Add
    Menu, MainMenu, Add, Exit, ExitScript
}

RebuildCustomPromptMenu(){
    global Prompts
    Menu, CustomPromptMenu, Add, __placeholder__, _
    Menu, CustomPromptMenu, DeleteAll
    if (Prompts.MaxIndex()){
        for index, prompt in Prompts
            Menu, CustomPromptMenu, Add, % prompt.Name, CustomPromptMenuHandler
    } else {
        Menu, CustomPromptMenu, Add, No prompts defined, _
    }
}

; =====================================================================================
; PREFERENCES / SETTINGS
; =====================================================================================
ShowPreferences:
    ShowPreferences_Fn()
return

ShowPreferences_Fn(){
    global Settings, DefaultEndpoints, Hotkeys, ColorPresets
    global PrefProvider, PrefApiKey, PrefEndpoint, PrefModel, PrefEpicNotes, PrefAITextColor, PrefColorPreview, PrefColorChoice
    global HK_Gen, HK_Err, HK_Restore, HK_Prompt, HK_Epic
    Gui, Pref:Destroy
    Gui, Pref:New, +AlwaysOnTop +OwnerMain, AI Tools Preferences
    Gui, Pref:Color, 121212, 1c1c1c
    Gui, Pref:Margin, 14, 12
    Gui, Pref:Font, s9 cFFFFFF, Segoe UI

    ; Store current color for the picker
    PrefAITextColor := Settings.AITextColor

    Gui, Pref:Add, GroupBox, xm ym w420 h250, API Settings
    Gui, Pref:Add, Text, xp+10 yp+20, Provider
    Gui, Pref:Add, DropDownList, vPrefProvider w260 gPrefProviderChanged Choose1, Azure OpenAI|OpenAI|Ollama
    GuiControl, Pref:ChooseString, PrefProvider, % Settings.Provider

    Gui, Pref:Add, Text, xm+10 yp+28, API key (blank for Ollama/local)
    Gui, Pref:Add, Edit, vPrefApiKey w260 Password, % Settings.ApiKey

    Gui, Pref:Add, Text, xm+10 yp+28, Endpoint / Base URL
    Gui, Pref:Add, Edit, vPrefEndpoint w360, % Settings.Endpoint

    Gui, Pref:Add, Text, xm+10 yp+28, Model name
    Gui, Pref:Add, Edit, vPrefModel w260, % Settings.Model

    epicFlag := Settings.EnableEpicNotes ? "Checked" : ""
    Gui, Pref:Add, Checkbox, xm+10 yp+30 %epicFlag% vPrefEpicNotes, Enable AI Notes (Epic copy/paste helper)

    ; AI Text Color presets
    Gui, Pref:Add, Text, xm+10 y+15, AI-Generated Text Color:
    colorBGR := Settings.AITextColor
    colorRGB := BGR2RGB(colorBGR)
    Gui, Pref:Add, Progress, x+10 yp-3 w40 h20 Background%colorRGB% vPrefColorPreview
    Gui, Pref:Add, DropDownList, x+10 yp-1 w180 gPrefColorChanged vPrefColorChoice, % BuildColorPresetList()
    Pref_SelectColorChoice(colorBGR)

    Gui, Pref:Add, GroupBox, xm y+16 w420 h180, Hotkeys (blank = disable)

    ; Row 1: Generate Impression and Check Errors
    Gui, Pref:Add, Text, xm+10 yp+20, Generate Impression:
    Gui, Pref:Add, Hotkey, vHK_Gen xm+10 y+5 w180, % Hotkeys.GenerateImpression
    Gui, Pref:Add, Text, xm+220 yp-20, Check Errors:
    Gui, Pref:Add, Hotkey, vHK_Err xm+220 y+5 w180, % Hotkeys.CheckErrors

    ; Row 2: Restore Dictation and Custom Prompt Menu
    Gui, Pref:Add, Text, xm+10 y+10, Restore Dictation:
    Gui, Pref:Add, Hotkey, vHK_Restore xm+10 y+5 w180, % Hotkeys.RestoreDictation
    Gui, Pref:Add, Text, xm+220 yp-20, Custom Prompt Menu:
    Gui, Pref:Add, Hotkey, vHK_Prompt xm+220 y+5 w180, % Hotkeys.CustomPrompt

    ; Row 3: Epic Notes
    Gui, Pref:Add, Text, xm+10 y+10, Epic Notes:
    Gui, Pref:Add, Hotkey, vHK_Epic xm+10 y+5 w180, % Hotkeys.EpicNotes

    Gui, Pref:Add, Button, xm y+22 w120 gSavePreferences, Save
    Gui, Pref:Add, Button, x+8 w140 gPrefRestoreDefaults, Restore Defaults
    Gui, Pref:Add, Button, x+8 w120 gPrefClose, Cancel
    Pref_EnableDisableKeyField(Settings.Provider)
    Gui, Pref:Show
}

PrefProviderChanged:
    PrefProviderChanged_Fn()
return

PrefProviderChanged_Fn(){
    global PrefProvider, PrefEndpoint, DefaultEndpoints
    Gui, Pref:Submit, NoHide
    Pref_EnableDisableKeyField(PrefProvider)
    if (DefaultEndpoints.HasKey(PrefProvider)){
        GuiControl, Pref:, PrefEndpoint, % DefaultEndpoints[PrefProvider]
    }
}

Pref_EnableDisableKeyField(provider){
    if (provider = "Ollama"){
        GuiControl, Pref:Disable, PrefApiKey
    } else {
        GuiControl, Pref:Enable, PrefApiKey
    }
}

BuildColorPresetList(){
    global ColorPresets
    list := ""
    for _, preset in ColorPresets
        list .= (list = "" ? "" : "|") . preset.Name
    return list
}

Pref_FindPresetIndex(bgr){
    global ColorPresets
    target := "0x" . NormalizeColor6(bgr)
    for i, preset in ColorPresets
        if ("0x" . NormalizeColor6(preset.BGR) = target)
            return i
    return 0
}

Pref_SelectColorChoice(bgr){
    global PrefAITextColor
    idx := Pref_FindPresetIndex(bgr)
    if (!idx)
        idx := 1
    GuiControl, Pref:Choose, PrefColorChoice, %idx%
    PrefAITextColor := bgr
    Pref_UpdateColorPreview(bgr)
}

Pref_GetSelectedPresetIndex(){
    global ColorPresets
    GuiControlGet, sel,, PrefColorChoice
    for i, preset in ColorPresets
        if (preset.Name = sel)
            return i
    return 0
}

Pref_UpdateColorPreview(bgr){
    colorRGB := BGR2RGB(bgr)
    GuiControl, Pref:+Background%colorRGB%, PrefColorPreview
}

PrefColorChanged:
    PrefColorChanged_Fn()
return

PrefColorChanged_Fn(){
    global ColorPresets, PrefAITextColor
    Gui, Pref:Submit, NoHide
    idx := Pref_GetSelectedPresetIndex()
    if (idx){
        PrefAITextColor := ColorPresets[idx].BGR
    }
    Pref_UpdateColorPreview(PrefAITextColor)
}

SavePreferences:
    SavePreferences_Fn()
return

SavePreferences_Fn(){
    global Settings, Hotkeys
    global PrefProvider, PrefApiKey, PrefEndpoint, PrefModel, PrefEpicNotes, PrefAITextColor
    global HK_Gen, HK_Err, HK_Restore, HK_Prompt, HK_Epic
    Gui, Pref:Submit
    Settings.Provider        := PrefProvider
    Settings.ApiKey          := PrefApiKey
    Settings.Endpoint        := PrefEndpoint
    Settings.Model           := PrefModel
    Settings.EnableEpicNotes := PrefEpicNotes ? 1 : 0
    Settings.AITextColor     := PrefAITextColor
    SaveSettings()

    Hotkeys.GenerateImpression := HK_Gen
    Hotkeys.CheckErrors        := HK_Err
    Hotkeys.RestoreDictation   := HK_Restore
    Hotkeys.CustomPrompt       := HK_Prompt
    Hotkeys.EpicNotes          := HK_Epic
    SaveHotkeys()
    RegisterAllHotkeys()

    UpdateStatusLine()
    Pref_Close_Fn()
}

PrefRestoreDefaults:
    PrefRestoreDefaults_Fn()
return

PrefRestoreDefaults_Fn(){
    global DefaultSettings, DefaultHotkeys, Settings, Hotkeys
    Settings := DefaultSettings.Clone()
    Hotkeys  := DefaultHotkeys.Clone()
    SaveSettings(), SaveHotkeys()
    Gui, Pref:Destroy
    ShowPreferences_Fn()
    UpdateStatusLine()
    RegisterAllHotkeys()
}

PickAIColor:
    PickAIColor_Fn()
return

PickAIColor_Fn(){
    ; Legacy stub kept for compatibility (color picker replaced with presets)
    return
}

PrefClose:
    Pref_Close_Fn()
return

Pref_Close_Fn(){
    Gui, Pref:Destroy
}

; =====================================================================================
; PROMPT MANAGER
; =====================================================================================
ShowPromptManager:
    ShowPromptManager_Fn()
return

ShowPromptManager_Fn(){
    global Prompts, PM_List
    Gui, PM:Destroy
    Gui, PM:New, +AlwaysOnTop +OwnerMain, Prompt Manager
    Gui, PM:Color, 121212, 1c1c1c
    Gui, PM:Margin, 10, 10
    Gui, PM:Font, cFFFFFF
    Gui, PM:Add, ListView, vPM_List w520 h220 Grid AltSubmit gPM_OnSelect, Name|Prompt|Hotkey
    for index, prompt in Prompts
        LV_Add("", prompt.Name, SubStr(prompt.Text,1,120) (StrLen(prompt.Text)>120 ? "..." : ""), prompt.Hotkey)
    Gui, PM:Add, Button, gPM_Add w120, Add
    Gui, PM:Add, Button, x+6 gPM_Edit w120, Edit
    Gui, PM:Add, Button, x+6 gPM_Delete w120, Delete
    Gui, PM:Add, Button, x+6 gPM_Close w120, Close
    Gui, PM:Show
}

PM_OnSelect:
return

PM_Add:
    PromptEditor()
return

PM_Edit:
    Row := LV_GetNext(0, "Focused")
    if (!Row){
        MsgBox, 48, Prompt Manager, Select a prompt to edit.
        return
    }
    LV_GetText(name, Row, 1)
    LV_GetText(_, Row, 2)
    PromptEditor(name)
return

PM_Delete:
    global Prompts
    Row := LV_GetNext(0, "Focused")
    if (!Row){
        MsgBox, 48, Prompt Manager, Select a prompt to delete.
        return
    }
    LV_GetText(name, Row, 1)
    for i, prompt in Prompts
        if (prompt.Name = name){
            Prompts.RemoveAt(i)
            break
        }
    SavePrompts()
    Gui, PM:Destroy
    RebuildCustomPromptMenu()
    ShowPromptManager_Fn()
return

PM_Close:
    Gui, PM:Destroy
return

PromptEditor(existingName:=""){
    global Prompts, PE_Name, PE_Text
    Gui, PE:Destroy
    Gui, PE:New, +OwnerPM +AlwaysOnTop, Edit Prompt
    Gui, PE:Color, 121212, 1c1c1c
    Gui, PE:Margin, 10, 10
    Gui, PE:Font, cFFFFFF
    if (existingName){
        for _, prompt in Prompts
            if (prompt.Name = existingName){
                current := prompt
                break
            }
    }
    Gui, PE:Add, Text,, Prompt name
    Gui, PE:Add, Edit, vPE_Name w360, % current.Name
    Gui, PE:Add, Text, y+8, Prompt text
    Gui, PE:Add, Edit, vPE_Text w360 h160, % current.Text
    Gui, PE:Add, Text, y+8, Hotkey (optional)
    Gui, PE:Add, Hotkey, vPE_Hotkey w200, % current.Hotkey
    Gui, PE:Add, Button, gPE_Save w120, Save
    Gui, PE:Add, Button, x+6 gPE_Cancel w120, Cancel
    Gui, PE:Show
}

PE_Save:
    global Prompts
    Gui, PE:Submit
    if (PE_Name = "" || PE_Text = ""){
        MsgBox, 48, Prompt Manager, Please provide both a name and prompt text.
        return
    }
    replaced := false
    for idx, prompt in Prompts {
        if (prompt.Name = PE_Name){
            Prompts[idx] := {"Name": PE_Name, "Text": PE_Text, "Hotkey": PE_Hotkey}
            replaced := true
            break
        }
    }
    if (!replaced)
        Prompts.Push({"Name": PE_Name, "Text": PE_Text, "Hotkey": PE_Hotkey})

    SavePrompts()
    Gui, PE:Destroy
    Gui, PM:Destroy
    RebuildCustomPromptMenu()
    ShowPromptManager_Fn()
return

PE_Cancel:
    Gui, PE:Destroy
return

; =====================================================================================
; HOTKEY REGISTRATION / HELP
; =====================================================================================
RegisterAllHotkeys(){
    global Hotkeys
    RegisterOneHotkey(Hotkeys.GenerateImpression, "GenerateImpression")
    RegisterOneHotkey(Hotkeys.CheckErrors, "CheckReportErrors")
    RegisterOneHotkey(Hotkeys.RestoreDictation, "RestoreDictation")
    RegisterOneHotkey(Hotkeys.CustomPrompt, "TriggerCustomPromptSelector")
    RegisterOneHotkey(Hotkeys.EpicNotes, "RunEpicNote")
    RegisterPromptHotkeys()
}

RegisterOneHotkey(combo, label){
    if (combo = "")
        return
    Hotkey, %combo%, %label%, On
}

RegisterPromptHotkeys(){
    global PromptHotkeyMap, Prompts
    ; clear existing prompt hotkeys
    for combo, _idx in PromptHotkeyMap
        Hotkey, %combo%, PromptHotkeyDispatcher, Off
    PromptHotkeyMap := {}
    for idx, prompt in Prompts {
        hk := prompt.HasKey("Hotkey") ? Trim(prompt.Hotkey) : ""
        ; Require a modifier (# ^ ! +) OR be a function key F1-F12. No plain words like "ERROR".
        validSyntax := RegExMatch(hk, "i)^[#^!+*~<>]*((F(1[0-2]?|[1-9]))|([A-Z0-9]))$")
        hasModifier := RegExMatch(hk, "[#^!+]")
        isFunction  := RegExMatch(hk, "i)^F(1[0-2]?|[1-9])$")
        if (hk != "" && validSyntax && (hasModifier || isFunction)){
            PromptHotkeyMap[hk] := idx
            Hotkey, %hk%, PromptHotkeyDispatcher, On
            Prompts[idx].Hotkey := hk  ; persist cleaned
        } else if (hk != "") {
            ; invalid hotkey string, drop it and persist cleanup
            Prompts[idx].Hotkey := ""
        }
    }
}

PromptHotkeyDispatcher:
    global PromptHotkeyMap
    idx := PromptHotkeyMap[A_ThisHotkey]
    if (idx)
        ApplyPrompt(idx)
return

ShowHotkeyHelp:
    ShowHotkeyHelp_Fn()
return

ShowHotkeyHelp_Fn(){
    global Hotkeys
    txt := "Configured hotkeys:`n"
    txt .= "Generate Impression: " . (Hotkeys.GenerateImpression ? Hotkeys.GenerateImpression : "None") . "`n"
    txt .= "Check Report for Errors: " . (Hotkeys.CheckErrors ? Hotkeys.CheckErrors : "None") . "`n"
    txt .= "Restore Dictation: " . (Hotkeys.RestoreDictation ? Hotkeys.RestoreDictation : "None") . "`n"
    txt .= "Apply Custom Prompt: " . (Hotkeys.CustomPrompt ? Hotkeys.CustomPrompt : "None") . "`n"
    txt .= "AI Notes (Epic): " . (Hotkeys.EpicNotes ? Hotkeys.EpicNotes : "None")
    MsgBox, 64, Hotkeys, %txt%
}

; =====================================================================================
; CORE GPT ACTIONS
; =====================================================================================
GenerateImpression:
    GenerateImpression_Fn()
return

GenerateImpression_Fn(){
    global pscribetxtbox, GPT_Impression_Numbered
    if (!EnsureApiConfig())
        return
    GetPowerscribeTextboxname()
    ControlGetText, PS_Report, %pscribetxtbox%, PowerScribe 360 | Reporting
    if (PS_Report = ""){
        MsgBox, 48, Generate Impression, Could not read report text. Click inside PowerScribe then retry.
        return
    }
    AddHistory("Generate Impression", PS_Report)

    ; Use multi-shot impression generation (will use examples if available, or zero-shot if not)
    rawImpr := OnGenerateMultiShotImpression(PS_Report)
    if (rawImpr = "")
        return
    needle := GPT_CleanImpression(rawImpr, GPT_Impression_Numbered)
    PowerscribeSendText2Target("IMPRESSION:", needle)
    TooltipTimed("Impression inserted.", 1200)
}

CheckReportErrors:
    CheckReportErrors_Fn()
return

CheckReportErrors_Fn(){
    global pscribetxtbox
    if (!EnsureApiConfig())
        return
    GetPowerscribeTextboxname()
    ControlGetText, PS_Report, %pscribetxtbox%, PowerScribe 360 | Reporting
    if (PS_Report = ""){
        MsgBox, 48, Check Errors, Could not read text from PowerScribe.
        return
    }
    AddHistory("Check Report Errors (pre-check)", PS_Report)

    ; Build the prompt with current date
    LLM_Prompt := "The current date is: " . A_MMMM . "-" . A_DD . "-" . A_YYYY
    LLM_Prompt .= ". Based on the radiology report provided, under the heading ERRORS:, create a list of any typos, voice transcription errors, conflicting statements, and grammar errors from the findings of the report. "
    LLM_Prompt .= "If there are no errors, under the heading ERRORS:, just output the word none. "
    LLM_Prompt .= "Aim to be concise while being precise. Focus on critical errors. Do not focus on minor grammatical errors. "
    LLM_Prompt .= "Do not focus on whether or not a statement is awkward. Findings can be repeated in the FINDINGS and IMPRESSION sections. "
    LLM_Prompt .= "Content in the IMPRESSION section do not need to be in full sentences.`n`n###`n`n"

    result := CallChatModel(LLM_Prompt, PS_Report, 0.0, 400)
    if (result = "")
        return
    GPT_GenerateErrorTerminalAPI(result)
}

RestoreDictation:
    RestoreDictation_Fn()
return

RestoreDictation_Fn(){
    global HistoryStack
    if (!HistoryStack.MaxIndex()){
        MsgBox, 48, Restore Dictation, Nothing to restore yet.
        return
    }
    prior := HistoryStack.RemoveAt(HistoryStack.MaxIndex())
    ReplaceEntireText(prior.Text)
    TooltipTimed("Restored: " prior.Label, 1500)
}

PullIndication:
    PullIndication_Fn()
return

PullIndication_Fn(){
    if (!EnsureApiConfig())
        return
    note := GetSelectedOrClipboardText()
    if (note = ""){
        MsgBox, 48, Pull Indication, Highlight text in Epic (or copy to clipboard) before running this.
        return
    }
    AddHistory("Pull Indication", GetActiveReportText())
    sysPrompt := "Condense the following clinical note into a clear, single-sentence radiology indication. Keep age, sex, symptoms, and relevant history. Remove headers and noise."
    cleaned := CallChatModel(sysPrompt, note, 0.2, 200)
    if (cleaned = "")
        return
    ReplaceSelection(cleaned)  ; if no selection, falls back to insert at caret
    TooltipTimed("Indication inserted.", 1500)
}

RunEpicNote:
    RunEpicNote_Fn()
return

RunEpicNote_Fn(){
    global Settings
    if (!Settings.EnableEpicNotes){
        MsgBox, 48, AI Notes, Enable AI Notes in Preferences to use this feature.
        return
    }
    if (!EnsureApiConfig())
        return
    note := GetSelectedOrClipboardText()
    if (note = ""){
        MsgBox, 48, AI Notes, Copy text from Epic (or another EMR window) before running this.
        return
    }
    sysPrompt := "You assist clinicians by producing concise, well-formatted notes. Summarize the following text, highlight key findings, and provide an optional short plan. Avoid PHI leakage beyond the provided text."
    summary := CallChatModel(sysPrompt, note, 0.2, 400)
    if (summary = "")
        return
    ShowEpicNoteWindow(note, summary)
}

TriggerCustomPromptSelector:
    TriggerCustomPromptSelector_Fn()
return

TriggerCustomPromptSelector_Fn(){
    global Prompts
    if (!Prompts.MaxIndex()){
        MsgBox, 48, Custom Prompts, No prompts defined. Use Prompt Manager to add some.
        return
    }
    Menu, CustomPromptMenu, Show
}

CustomPromptMenuHandler:
    CustomPromptMenuHandler_Fn()
return

CustomPromptMenuHandler_Fn(){
    ApplyPromptByName(A_ThisMenuItem)
}

ApplyPromptByName(name){
    global Prompts
    for idx, prompt in Prompts
        if (prompt.Name = name){
            ApplyPrompt(idx)
            return
        }
    MsgBox, 48, Custom Prompt, Prompt not found.
}

ApplyPromptByIndex(idx){
    global Prompts
    if (idx < 1 || idx > Prompts.MaxIndex()){
        MsgBox, 48, Custom Prompt, Invalid prompt number.
        return
    }
    ApplyPrompt(idx)
}

ApplyPrompt(idx){
    global Prompts, Settings
    if (!EnsureApiConfig())
        return
    prompt := Prompts[idx]
    text := GetSelectionOrReport()
    if (text = ""){
        MsgBox, 48, Custom Prompt, Highlight text first (or place the caret in a report).
        return
    }
    AddHistory("Custom Prompt - " prompt.Name, GetActiveReportText())
    sysPrompt := "Apply the following instruction to the provided text. Preserve clinical facts and measurements. If the instruction is to rewrite, keep the meaning faithful."
    userText := "Instruction: " . prompt.Text . "`n---`nText:`n" . text
    result := CallChatModel(sysPrompt, userText, 0.2, 600)
    if (result = "")
        return
    ReplaceSelection(result, Settings.AITextColor)
    TooltipTimed("Applied prompt: " prompt.Name, 1500)
}

; =====================================================================================
; FEW-SHOT REPORT ARCHIVE + PROMPT BUILDER (cJson)
; =====================================================================================
AddReport(reportText, examType)
{
    global archiveDir, maxExamplesPerExam, PromptMessages

    ; strip common signature line
    sigPos := InStr(reportText, "This report electronically signed by")
    if (sigPos){
        lines := StrSplit(reportText, "`n")
        cleaned := []
        for _, ln in lines
            if (!InStr(ln, "This report electronically signed by"))
                cleaned.Push(ln)
        reportText := Join("`n", cleaned*)
    }

    SplitReport(reportText, preImpr, impr)
    if (preImpr = "" || impr = ""){
        Tooltip, IMPRESSION delimiter not found.
        return false
    }

    userBlock      := Trim(preImpr)
    assistantBlock := Trim(impr)
    jsonLine       := JSON.Dump({ "user": userBlock , "assistant": assistantBlock }, 0)

    safeName := SafeFile(examType)
    filePath := archiveDir . "\\" . safeName . ".jsonl"

    if (FileExist(filePath)){
        curr := LoadAllLines(filePath)
        if (curr.Length() >= maxExamplesPerExam){
            Tooltip, % "Already have " curr.Length() " examples (max = " maxExamplesPerExam ")."
            return false
        }
        for each, ln in curr
            if (ln = jsonLine){
                Tooltip, This report already exists.
                return false
            }
    }

    FileAppend, % jsonLine . "`n", % filePath
    Tooltip, Report added successfully.
    SetTimer, __HideTip, -2000
    PromptMessages[safeName] := BuildFewShotMessages(safeName)
    return true
}

GetExamples(examType, n := 5) {
    global archiveDir
    filePath := archiveDir . "\\" . examType . ".jsonl"
    if (!FileExist(filePath))
        return []
    FileRead, text, %filePath%
    arr := []
    Loop, Parse, text, `n, `r
        if (A_LoopField != "")
            arr.Push(A_LoopField)

    start := Max(1, arr.Length() - n + 1)
    examples := []
    for i, ln in arr
        if (i >= start)
            examples.Push(ln)
    return examples
}

BuildFewShotMessages(examType, n := 5) {
    examples := GetExamples(examType, n)
    messages := []
    for each, line in examples {
        cleaned := NormalizeJsonLine(line)
        if (cleaned = "")
            continue
        try obj := JSON.Load(cleaned)
        catch {
            continue
        }
        if (!IsObject(obj) || !obj.HasKey("user") || !obj.HasKey("assistant"))
            continue
        messages.Push({"role": "user", "content": obj["user"]})
        messages.Push({"role": "assistant", "content": obj["assistant"]})
    }
    return messages
}

BuildPromptMap(nExamples := 5)
{
    promptMap := {}
    types := ListExamTypes()
    for each, examType in types {
        msgs := BuildFewShotMessages(examType, nExamples)
        promptMap[examType] := msgs
    }
    return promptMap
}

NormalizeJsonLine(line){
    ; Replace common smart quotes/control chars so JSON.Load won't choke on legacy lines.
    if (line = "")
        return ""
    line := RegExReplace(line, "^\xEF\xBB\xBF")             ; strip UTF-8 BOM if present
    line := StrReplace(line, Chr(0x201C), """")             ; left double quote
    line := StrReplace(line, Chr(0x201D), """")             ; right double quote
    line := StrReplace(line, Chr(0x2018), "'")              ; left single quote
    line := StrReplace(line, Chr(0x2019), "'")              ; right single quote
    line := RegExReplace(line, "[\x00-\x08\x0B-\x0C\x0E-\x1F]") ; strip control chars
    return Trim(line)
}

SplitReport(fullText, ByRef beforeImpr, ByRef afterImpr) {
    SplitPos := InStr(fullText, "IMPRESSION:", false)
    if (!SplitPos) {
        beforeImpr := afterImpr := ""
        return
    }
    SplitPos += StrLen("IMPRESSION:")
    beforeImpr := SubStr(fullText, 1, SplitPos-1)
    afterImpr  := SubStr(fullText, SplitPos)
}

LoadAllLines(path) {
    lines := []
    if (FileExist(path)) {
        FileRead, text, % path
        Loop, Parse, text, `n, `r
            if (A_LoopField != "")
                lines.Push(A_LoopField)
    }
    return lines
}

SaveAllLines(path, ByRef lines) {
    fh := FileOpen(path, "w", "UTF-8")
    if !IsObject(fh) {
        MsgBox 16, I/O error, % "Couldn’t open " path
        return
    }
    for _, ln in lines
        fh.Write(ln "`n")
    fh.Close()
}

SafeFile(examType) {
    safe := RegExReplace(examType, "[\\/:*?""<>|]", "-")
    safe := RegExReplace(safe, "[\.\s]+$", "")
    return safe
}

OnGenerateMultiShotImpression(Report){
	Global Settings

	temperature := 0
	maxTokens  := 1000
	messages := BuildPrompt(Report)
	if (!IsObject(messages))
		return ""

	try {
		payload := {"messages": messages, "temperature": temperature, "model": Settings.Model}

		; Add max_tokens unless using Ollama
		if (Settings.Provider != "Ollama")
			payload["max_tokens"] := maxTokens

		; For Ollama, disable streaming
		if (Settings.Provider = "Ollama")
			payload["stream"] := false

		jsonBody := Jxon_Dump(payload)

		http := ComObjCreate("WinHttp.WinHttpRequest.5.1")
		http.Open("POST", Settings.Endpoint, false)
		http.SetTimeouts(5000, 5000, 30000, 30000)
		http.SetRequestHeader("Content-Type","application/json")

		; Set appropriate auth header based on provider
		if (Settings.Provider = "OpenAI")
			http.SetRequestHeader("Authorization","Bearer " Settings.ApiKey)
		else if (Settings.Provider = "Azure OpenAI")
			http.SetRequestHeader("api-key", Settings.ApiKey)

		http.Send(jsonBody)
		http.WaitForResponse()
		status := http.Status
		resp := http.ResponseText
	} catch e {
		MsgBox, 16, API Error, % "Failed to call endpoint.`n" e.Message
		return ""
	}

	if (status < 200 || status >= 300){
		MsgBox, 16, API Error, % "HTTP " status "`n" resp
		return ""
	}

	; Parse response (works for both OpenAI and Ollama)
	respObj := Jxon_Load(resp)
	if (!IsObject(respObj))
		return ""

	if (Settings.Provider = "Ollama"){
		try return respObj["message"]["content"]
	} else {
		try return respObj["choices"][1]["message"]["content"]
	}
	return ""
}

OnDeleteExample()
{
    Gui, FSP:Default
    Gui, Submit, NoHide
    global LBExam, LBIdx, archiveDir, gExamples, PromptMessages

    if (LBExam = "" || LBIdx = "")
    {
        MsgBox, 48, Nothing selected, Select an exam type and example first.
        return
    }

    filePath   := archiveDir . "\\" . LBExam . ".jsonl"
    allLines   := LoadAllLines(filePath)
    targetLine := gExamples[LBIdx]

    MsgBox, 4, Confirm, Delete example #%LBIdx% for %LBExam%?
    IfMsgBox, No
        return

	for idx, ln in allLines
		if (ln = targetLine) {
			allLines.RemoveAt(idx)
			break
		}

    SaveAllLines(filePath, allLines)
    PromptMessages[LBExam] := BuildFewShotMessages(LBExam)
    MsgBox, 64, Deleted, Example removed.
    OnExamSelect()
}

OnAddExam(){
    GetPowerscribeTextboxname()
    Global pscribetxtbox
    Controlgettext, PS_Report, %pscribetxtbox%, PowerScribe 360 | Reporting
	
	if(PS_Report == ""){
		Msgbox, No report is selected
		Return
	}

	WinGet, MMX2, MinMax, PowerScribe 360 | Reporting
	If (MMX2 == -1)
		WinRestore, PowerScribe 360 | Reporting
	WinGetText, panel, PowerScribe 360 | Reporting
	
	if((InStr(panel, "Report - ")) || (InStr(panel, "Addendum - "))) {
		OpenPowerscribeSidePanel()
		ExamName :=  GetPowerscribeElementInfo("Description:")
	} else {
		if(!InStr(panel, ".) - ")){
			ControlGetPos, px, py, pw, ph,WindowsForms10.Window.8.app.0.26ac0ad_r8_ad16, PowerScribe 360 | Reporting
			btnX := px + 2
			btnY := py + 15
			ControlClick, x%btnX% y%btnY%, PowerScribe 360 | Reporting, , , 1, NA
		}
		ExamName := GetPowerscribeElementInfo("Exam Date:")
	}
	
    AddReport(PS_Report, ExamName)
	OnExamSelect() 
	UpdateExamList()
}

OnBulkAddExam:
    OnBulkAddExam_Fn()
return

OnBulkAddExam_Fn(maxBatch := 500, pollMS := 250, timeoutMS := 8000)
{
    Global pscribetxtbox
    listCtl  := "WindowsForms10.Window.8.app.0.26ac0ad_r8_ad14"
    winTitle := "PowerScribe 360 | Reporting"

    ; Activate the PowerScribe window
    WinActivate, %winTitle%
    if ErrorLevel {
        MsgBox, 48, Bulk Add, Could not activate PowerScribe window.
        return
    }
    Sleep, 200

    ; Get the position of the list control so we can click on it
    ControlGetPos, listX, listY, listW, listH, %listCtl%, %winTitle%
    if ErrorLevel {
        MsgBox, 48, Bulk Add, Could not find report list control.
        return
    }

    ; Click in the middle of the list control to ensure it has focus
    clickX := listX + (listW // 2)
    clickY := listY + 30  ; Click near the top but not on the header
    ControlClick, x%clickX% y%clickY%, %winTitle%, , LEFT, 1, NA
    Sleep, 200

    ; Go to top of the report list using keyboard
    Send {Home}
    Sleep, 300
    Send {Home}  ; Send twice to ensure we're at the top
    Sleep, 500

    ; Ensure we have the PowerScribe textbox identified
    GetPowerscribeTextboxname()
    if (pscribetxtbox = "") {
        MsgBox, 48, Bulk Add, Could not identify PowerScribe text box.
        return
    }

    ; Get the initial report text (first report in list)
    Sleep, 300
    ControlGetText, PriorReport, %pscribetxtbox%, %winTitle%

    ; If the first report is empty, we're not in a valid state
    if (PriorReport = "") {
        MsgBox, 48, Bulk Add, No report selected. Please select a report first.
        return
    }

    loopCount := 0

    Loop, %maxBatch% {
        ; Try to add the current report
        OnAddExam()
        loopCount++

        ; Now move down to the next report
        WinActivate, %winTitle%
        Sleep, 100

        ; Click on the list to ensure focus, then send Down
        ControlClick, x%clickX% y%clickY%, %winTitle%, , LEFT, 1, NA
        Sleep, 100
        Send {Down}
        Sleep, 200

        ; Wait for the report text to change
        elapsed := 0
        CurrentReport := PriorReport
        while (CurrentReport = PriorReport && elapsed < timeoutMS) {
            Sleep, %pollMS%
            elapsed += pollMS
            ControlGetText, CurrentReport, %pscribetxtbox%, %winTitle%
        }

        ; Check if we've reached the end of the list
        if (CurrentReport = PriorReport) {
            MsgBox, 64, Bulk Add, End of list reached. Processed %loopCount% report(s).
            break
        }

        ; Check if the new report is empty (shouldn't happen, but safety check)
        if (CurrentReport = "") {
            MsgBox, 48, Bulk Add, Empty report encountered. Processed %loopCount% report(s).
            break
        }

        ; Update for next iteration
        PriorReport := CurrentReport
    }

    if (loopCount > 0)
        MsgBox, 64, Bulk Add, Successfully processed %loopCount% report(s).
    else
        MsgBox, 48, Bulk Add, No reports were added.
}

OnExamSelect() {
    Gui, FSP:Default
    Gui, Submit, NoHide
    global LBExam, LBIdx, EditExamples, BtnDelete, gExamples

    GuiControl,, EditExamples,
    GuiControl,, LBIdx, |
    GuiControl, Disable, BtnDelete

    if (LBExam = "")
        return

    gExamples := GetExamples(LBExam, maxExamplesPerExam)
	
    if (gExamples.Length()) {
        idxList := "|"
        Loop, % gExamples.Length(){
			idxList .= A_Index . "|"
		}
        GuiControl,, LBIdx, % idxList
        GuiControl, Show, LBIdx
    } else {
        GuiControl,, EditExamples,
    }
}

OnExampleIndexSelect() {
    Gui, FSP:Default
    Gui, Submit, NoHide
    global LBIdx, gExamples, EditExamples, BtnDelete

    if (LBIdx = "") {
        GuiControl,, EditExamples,
        GuiControl, Disable, BtnDelete
        return
    }

    idx := LBIdx
    if (idx >= 1 && idx <= gExamples.Length()) {
        obj  := JSON.Load(gExamples[idx])
        text :=  obj["user"] .  obj["assistant"]
        GuiControl,, EditExamples, % text
        GuiControl, Enable, BtnDelete
    }
}

BuildPrompt(Report)  {
    Gui, Submit, NoHide
    global pscribetxtbox, PromptMessages, Settings

    ExamName := SafeFile(GetPowerscribeElementInfo("Description:"))
    if (ExamName = "")  {
        MsgBox, 48, Missing exam type, Please select an exam type first.
        return
    }

    if (PromptMessages.HasKey(ExamName))
        messages := ObjClone(PromptMessages[ExamName])
    else
        messages := []

    hasExamples := messages.Length()

    if (hasExamples) {
        sysText := "You are a radiology assistant tasked with generating concise radiology impressions in bullet points based on provided findings. Focus on clinically significant findings and prioritize them in order of importance. Avoid trivial or non-pertinent details. Study the style of prior impressions in the conversation history to replicate tone and structure."
    } else {
        sysText := "Under the heading IMPRESSION:, write a concise 1-5 bullet point radiology impression (combining any relevant statements into a single point) based on the findings. Arrange the bullet points in order of importance."
    }

    sysMsg := Object("role", "system", "content", sysText)
    messages.InsertAt(1, sysMsg)

    SplitReport(Report, preImpr, impr)
    if (preImpr = "")  {
        MsgBox, 16, Archive Error, IMPRESSION: delimiter not found. Please ensure your report contains the exact header IMPRESSION:
        return false
    }
    userBlock := Trim(preImpr)
    liveCase  := Object("role", "user", "content", userBlock)
    messages.Push(liveCase)
    return messages
}

ListExamTypes() {
    global archiveDir
    arr := []
    Loop, Files, % archiveDir "\\*.jsonl"
        arr.Push(RegExReplace(A_LoopFileName, "\.jsonl$"))
    return arr
}

Join(delim, ByRef input*) {
    out := ""
    for idx, val in input
        out .= (idx>1 ? delim : "") val
    return out
}

ObjClone(obj) {
    ; Deep clone an AHK object (array or associative array)
    if (!IsObject(obj))
        return obj

    clone := obj.Clone()
    for key, val in clone
        if (IsObject(val))
            clone[key] := ObjClone(val)
    return clone
}

UpdateExamList(){
    types := ListExamTypes()
    if (types.Length()) {
        listStr := "|" . Join("|", types*)
        GuiControl,, LBExam, % listStr
    }
}

; -------------------- FEW-SHOT BUILDER GUI -----------------------------

ShowFewShotBuilder(){
    global LBExam, LBIdx, EditExamples, BtnDelete, gExamples, MainbuttonH, MainbuttonGap, AppName
    Gui, FSP:Destroy
    WinGetPos, X, Y, Width, Height, %AppName%
    if (X = "")
        X := 100, Y := 100
    Gui, FSP:New, +AlwaysOnTop +OwnerMain, Few-Shot Prompt Builder
    Gui, FSP:Color, 121212, 1c1c1c
    Gui, FSP:Margin, 10, 10
    Gui, FSP:Font, cFFFFFF
    Gui, FSP:Add, Text, x10 y10 w200, Exam Type:
    Gui, FSP:Add, ListBox, xp y+5 w300 h200 vLBExam gOnExamSelect
    Gui, FSP:Add, Text, x+15 y10 w90, Example #:
    Gui, FSP:Add, ListBox, xp y+5 w90 h200 vLBIdx gOnExampleIndexSelect

    Gui, FSP:Add, Text, x10 y+12, Selected Example:
    Gui, FSP:Add, Edit, xp y+5 w520 h200 vEditExamples ReadOnly

    btnW := 150
    Gui, FSP:Add, Button, x10 y+12 w%btnW% gOnAddExam, Add Exam
    Gui, FSP:Add, Button, x+%MainbuttonGap% yp w%btnW% gOnDeleteExample vBtnDelete Disabled, Delete Example
    Gui, FSP:Add, Button, x+%MainbuttonGap% yp w%btnW% gOnBulkAddExam, Bulk Add Exams

    Gui, FSP:Show, x%X% y%Y%, Few-Shot Prompt Builder
    UpdateExamList()
    return
}

OpenFewShotBuilder:
	ShowFewShotBuilder()
return
; =====================================================================================
; ERROR TERMINAL (clickable lines with quoted text navigation)
; =====================================================================================
GPT_GenerateErrorTerminalAPI(paragraph){
	Global ErrCtrlMap, AppName, GPTToggleButton, quotedTextArray
	quotedTextArray := []
	clickVars := []

	; clean the text so it's ready to process and for output
	paragraph := RegExReplace(paragraph, "\.{2,}", "")

	; Normalize curly quotes
	LeftCurlyQuote := Chr(0x201C)   ; “
	RightCurlyQuote := Chr(0x201D)  ; ”
	paragraph := StrReplace(paragraph, LeftCurlyQuote, """""")
	paragraph := StrReplace(paragraph, RightCurlyQuote, """""")

	paragraph := RegExReplace(paragraph, "â", "-")

	pos := 1
	match := 0
	while (pos := RegExMatch(paragraph, """([^""]+)""", match, pos + StrLen(match))) {
		quotedTextArray.push(match1)
	}

	characterwrap := 100
	paragraphs := StrSplit(paragraph, "`n")
	lines := []

	; Split long lines so they wrap nicely
	for index, para in paragraphs {
		while (StrLen(para) > 0) {
			if (StrLen(para) > characterwrap) {
				part := SubStr(para, 1, characterwrap)
				position := InStr(part, " ", 0, -1)
				if (!position)
					position := characterwrap
				lines.Push(SubStr(para, 1, position))
				para := SubStr(para, position + 1)
			} else {
				para := para . "`n"
				lines.Push(para)
				break
			}
		}
	}

	; anchor position near existing terminal or main window
	errX := ""
	errY := ""
	if WinExist("ErrorTerminal"){
		WinGetPos errX, errY, errWidth, errHeight, ErrorTerminal
	} else if WinExist(AppName) {
		WinGetPos errX, errY, errWidth, errHeight, %AppName%
		errY := errY + errHeight + 5
	}
	errX := errX + 0
	errY := errY + 0
	if (errX = 0 && errY = 0){
		errX := 100
		errY := 100
	}

	PatientMRN := GetPowerscribeElementInfo("MRN:")
	ExamType := GetPowerscribeElementInfo("Description:")

	OnMessage(0x0201, "WM_LBUTTONDOWN")

	Gui, ErrorTerminal: New, +AlwaysOnTop, Error Report
	Gui, ErrorTerminal:Color, 121212, 1c1c1c
	Gui, ErrorTerminal:Margin, 12, 12
	Gui, ErrorTerminal:Font, s10 cFFFFFF, Segoe UI

	; Header
	Gui, ErrorTerminal:Font, s11 bold
	Gui, ErrorTerminal:Add, Text, x12 y12 w600, Report Error Check
	Gui, ErrorTerminal:Font, s9 norm

	; Exam info
	Gui, ErrorTerminal:Font, s9 cFFD700
	Gui, ErrorTerminal:Add, Text, x12 y+10, Exam Type:
	Gui, ErrorTerminal:Font, s9 cFFFFFF
	Gui, ErrorTerminal:Add, Text, x+5 yp, %ExamType%
	Gui, ErrorTerminal:Font, s9 cFFD700
	Gui, ErrorTerminal:Add, Text, x12 y+5, MRN:
	Gui, ErrorTerminal:Font, s9 cFFFFFF
	Gui, ErrorTerminal:Add, Text, x+5 yp, %PatientMRN%

	; Separator
	Gui, ErrorTerminal:Add, Text, x12 y+10 w600 h1 0x10

	; Error content
	Gui, ErrorTerminal:Font, s9 cFFFFFF
	Gui, ErrorTerminal:Add, Text, x12 y+10,

	totalquoteCounter := 0
	quoteoddeven := 0

		for index, line in lines {
			quoteoddeven := Mod(totalquoteCounter, 2)
			splitline := StrSplit(line, """""")

			Loop % splitline.MaxIndex(){
				if(A_INDEX == 1 && quoteoddeven == 1 && A_INDEX == splitline.MaxIndex()){
					temptext := splitline[A_INDEX]
					ctrlVar := "err" A_INDEX "_" totalquoteCounter
					clickVars.Push(ctrlVar)
					Gui, ErrorTerminal:Add, Text, c00D7FF x12 y+1 v%ctrlVar% gSearchErrorText, % temptext
				} else if(A_INDEX == 1 && quoteoddeven == 1){
					totalquoteCounter++
					temptext := splitline[A_INDEX] . """"""
					ctrlVar := "err" A_INDEX "_" totalquoteCounter
					clickVars.Push(ctrlVar)
					Gui, ErrorTerminal:Add, Text, c00D7FF x12 y+1 v%ctrlVar% gSearchErrorText, % temptext
				} else if (A_INDEX == 1){
					Gui, ErrorTerminal:Add, Text, x12 y+1, % splitline[A_INDEX]
				} else if ((Mod(A_INDEX, 2) == 0) && (A_INDEX == splitline.MaxIndex()) && (quoteoddeven == 0)){
					totalquoteCounter++
					temptext := """""" . splitline[A_INDEX]
					ctrlVar := "err" A_INDEX "_" totalquoteCounter
					clickVars.Push(ctrlVar)
					Gui, ErrorTerminal:Add, Text, c00D7FF v%ctrlVar% gSearchErrorText x+1 yp, %temptext%
				} else if ((Mod(A_INDEX, 2) == 1) && (A_INDEX == splitline.MaxIndex()) && (quoteoddeven == 1)){
					totalquoteCounter++
					temptext := """""" . splitline[A_INDEX]
					ctrlVar := "err" A_INDEX "_" totalquoteCounter
					clickVars.Push(ctrlVar)
					Gui, ErrorTerminal:Add, Text, c00D7FF v%ctrlVar% gSearchErrorText x+1 yp, %temptext%
				} else if (Mod(A_INDEX, 2) == 1 && quoteoddeven == 1) {
					totalquoteCounter := totalquoteCounter + 2
					temptext := splitline[A_INDEX]
					ctrlVar := "err" A_INDEX "_" totalquoteCounter
					clickVars.Push(ctrlVar)
					Gui, ErrorTerminal:Add, Text, c00D7FF v%ctrlVar% gSearchErrorText x+1 yp, "%temptext%"
				} else if (Mod(A_INDEX, 2) == 1){
					Gui, ErrorTerminal:Add, Text, x+1 yp, % splitline[A_INDEX]
				} else if (Mod(A_Index, 2) == 0 && quoteoddeven == 0){
					totalquoteCounter := totalquoteCounter + 2
					temptext := splitline[A_Index]
					ctrlVar := "err" A_Index "_" totalquoteCounter
					clickVars.Push(ctrlVar)
					Gui, ErrorTerminal:Add, Text, c00D7FF v%ctrlVar% gSearchErrorText x+1 yp, "%temptext%"
			} else if (Mod(A_Index, 2) == 0 && quoteoddeven == 1){
				Gui, ErrorTerminal:Add, Text, x+1 yp, % splitline[A_Index]
			}
		}
	}

	; register clickable controls so dragging logic ignores them
	ErrCtrlMap := {}
	for _, ctrlVar in clickVars {
		GuiControlGet, ctrlHwnd, ErrorTerminal:HWND, %ctrlVar%
		if (ctrlHwnd)
			ErrCtrlMap[ctrlHwnd] := 1
	}

	; Add close button at the bottom
	Gui, ErrorTerminal:Add, Text, x12 y+20 w600 h1 0x10
	Gui, ErrorTerminal:Font, s9
	Gui, ErrorTerminal:Add, Button, x12 y+10 w120 gExitErrorTerminal, Close

	Gui, ErrorTerminal: Show, x%errX% y%errY% AutoSize, Error Report
	WinGetPos, , , InitialWidth, InitialHeight, ErrorTerminal
	ErrorTerminalExpanded := true

	ErrorTerminalZeroed := false

	Return
}

; Function to toggle between collapsed and expanded states
ToggleErrorTerminal:
    if (ErrorTerminalExpanded := !ErrorTerminalExpanded) {
        ; Expand
        Gui, ErrorTerminal: Show, h%InitialHeight% w%InitialWidth% ; Use the stored initial height
        GuiControl, ErrorTerminal:Text, GPTToggleButton, Collapse
    } else {
        ; Collapse to show only header and exam info
        Gui, ErrorTerminal: Show, w%InitialWidth% h150
        GuiControl, ErrorTerminal:Text, GPTToggleButton, Expand
    }
return

SearchErrorText:
	global quotedTextArray
	temp_quote_number := strsplit(A_GuiControl, "_")[2]
	temp_quote_index := Ceil(temp_quote_number/2)
	searchstring := RegExReplace(quotedTextArray[temp_quote_index], "[,.]$", "")
	FindReportText(searchstring)
	;tooltip % A_GuiControl . " " . temp_quote_index .  " " . searchstring
	;Sleep 1000
	;tooltip
Return

RemoveToolTip:
    ToolTip
    SetTimer, RemoveToolTip, Off
Return

FindReportText(needle) {
    global pscribehwnd

    ; Get the text from the edit box
    text := RegExReplace(Edit_GetText(pscribehwnd, -1), "`r`n", "`n")  

    ; Get the current selection position
    pos := Edit_GetSel(pscribehwnd)

    ; Get the currently selected text
    currentSelText := SubStr(text, pos + 1, StrLen(needle))

    ; Determine the starting position for the search
    if (currentSelText = needle) {
        ; Start the search from the current selection position + 1
        searchPos := InStr(SubStr(text, pos + StrLen(needle) + 1), needle)
        if (searchPos > 0) {
            ; Update the position to the absolute position in the text
            searchPos := searchPos + pos + StrLen(needle)
        }
    } else {
        ; Start the search from the beginning of the text
        searchPos := InStr(text, needle)
    }

    ; If needle is found, highlight it
    if (searchPos > 0) {
        ; Highlight the found needle
        Edit_SetSel(pscribehwnd, searchPos - 1, searchPos + StrLen(needle) - 1)
		Winactivate, PowerScribe 360 | Reporting 
    } else {
        ; If not found, reset the selection to the start of the text
        Edit_SetSel(pscribehwnd, 0, 0)
    }
}

; =====================================================================================
; EDIT HELPER FUNCTIONS (simplified from Edit.ahk library)
; =====================================================================================
Edit_GetText(hEdit, p_Length=-1) {
    Static WM_GETTEXT := 0xD
    if (p_Length < 0)
        p_Length := Edit_GetTextLength(hEdit)

    VarSetCapacity(l_Text, p_Length * (A_IsUnicode ? 2 : 1) + 1, 0)
    SendMessage, WM_GETTEXT, p_Length + 1, &l_Text,, ahk_id %hEdit%
    Return l_Text
}

Edit_GetTextLength(hEdit) {
    Static WM_GETTEXTLENGTH := 0xE
    SendMessage, WM_GETTEXTLENGTH, 0, 0,, ahk_id %hEdit%
    Return ErrorLevel
}

Edit_GetSel(hEdit, ByRef r_StartSelPos="", ByRef r_EndSelPos="") {
    Static Dummy3304
          ,s_StartSelPos
          ,s_EndSelPos
          ,Dummy1 := VarSetCapacity(s_StartSelPos, 4, 0)
          ,Dummy2 := VarSetCapacity(s_EndSelPos, 4, 0)
          ,EM_GETSEL := 0xB0

    ; Get the select positions
    SendMessage, EM_GETSEL, &s_StartSelPos, &s_EndSelPos,, ahk_id %hEdit%
    r_StartSelPos := NumGet(s_StartSelPos, 0, "UInt")
    r_EndSelPos := NumGet(s_EndSelPos, 0, "UInt")
    Return r_StartSelPos
}

Edit_SetSel(hEdit, p_StartSelPos=0, p_EndSelPos=-1) {
    Static EM_SETSEL := 0xB1
    SendMessage, EM_SETSEL, p_StartSelPos, p_EndSelPos,, ahk_id %hEdit%
}

SetTextColorSafe(hEdit, startPos, endPos, colorBGR) {
    ; Safely apply color to a text range in a RichEdit control (cross-process safe)
    Static EM_SETCHARFORMAT := 0x0444
    Static EM_SETSEL := 0x00B1
    Static SCF_SELECTION := 0x0001
    Static CFM_COLOR := 0x40000000
    Static MEM_COMMIT := 0x1000, MEM_RESERVE := 0x2000, MEM_RELEASE := 0x8000, PAGE_READWRITE := 0x04
    Static PROCESS_ALL_ACCESS := 0x1F0FFF

    ; Validate positions
    if (startPos < 0 || endPos <= startPos)
        return

    ; Clamp to the current text length
    textLen := Edit_GetTextLength(hEdit)
    if (endPos > textLen)
        endPos := textLen
    if (endPos <= startPos)
        return

    ; Normalize and convert color to integer (expects BGR COLORREF)
    colorBGR := "0x" . NormalizeColor6(colorBGR)
    colorBGR += 0

    ; Select the text range using EM_SETSEL (wParam/lParam are cross-process safe)
    SendMessage, EM_SETSEL, startPos, endPos,, ahk_id %hEdit%
    Sleep, 20

    ; Prepare CHARFORMAT2 structure locally
    VarSetCapacity(CHARFORMAT2, 116, 0)
    NumPut(116, CHARFORMAT2, 0, "UInt")           ; cbSize
    NumPut(CFM_COLOR, CHARFORMAT2, 4, "UInt")     ; dwMask
    NumPut(0, CHARFORMAT2, 8, "UInt")             ; dwEffects (ensure CFE_AUTOCOLOR is off)
    NumPut(colorBGR, CHARFORMAT2, 20, "UInt")     ; crTextColor (BGR format)

    ; If the control lives in another process, marshal the struct into that process
    targetPid := 0
    DllCall("GetWindowThreadProcessId", "Ptr", hEdit, "UInt*", targetPid)
    if (!targetPid)
        return

    fmtPtr := &CHARFORMAT2
    hProc := 0, remotePtr := 0

    if (targetPid != DllCall("GetCurrentProcessId")) {
        hProc := DllCall("OpenProcess", "UInt", PROCESS_ALL_ACCESS, "Int", false, "UInt", targetPid, "Ptr")
        if (!hProc)
            return
        remotePtr := DllCall("VirtualAllocEx", "Ptr", hProc, "Ptr", 0, "UPtr", 116, "UInt", MEM_COMMIT|MEM_RESERVE, "UInt", PAGE_READWRITE, "Ptr")
        if (!remotePtr){
            DllCall("CloseHandle", "Ptr", hProc)
            return
        }
        DllCall("WriteProcessMemory", "Ptr", hProc, "Ptr", remotePtr, "Ptr", &CHARFORMAT2, "UPtr", 116, "UPtr*", 0)
        fmtPtr := remotePtr
    }

    ; Apply the color format
    SendMessage, EM_SETCHARFORMAT, SCF_SELECTION, fmtPtr,, ahk_id %hEdit%

    ; Cleanup remote allocation if used
    if (remotePtr){
        DllCall("VirtualFreeEx", "Ptr", hProc, "Ptr", remotePtr, "UPtr", 0, "UInt", MEM_RELEASE)
        DllCall("CloseHandle", "Ptr", hProc)
    }
}

ExitErrorTerminal:
	Gui, ErrorTerminal:Destroy
Return

; =====================================================================================
; EPIC NOTE RENDERER
; =====================================================================================
ShowEpicNoteWindow(original, summary){
    global EN_Output
    Gui, EN:Destroy
    Gui, EN:New, +AlwaysOnTop +OwnerMain, AI Notes
    Gui, EN:Color, 121212, 1c1c1c
    Gui, EN:Margin, 10, 10
    Gui, EN:Font, s9 cFFFFFF, Segoe UI
    Gui, EN:Add, Text,, Original selection (read-only):
    Gui, EN:Add, Edit, w420 h100 ReadOnly, % original
    Gui, EN:Add, Text, y+6, AI output:
    Gui, EN:Add, Edit, vEN_Output w420 h120, % summary
    Gui, EN:Add, Button, gEN_Copy w100, Copy
    Gui, EN:Add, Button, x+6 gEN_Insert w100, Insert into active window
    Gui, EN:Add, Button, x+6 gEN_Close w100, Close
    Gui, EN:Show
}

EN_Copy:
    Gui, EN:Submit, NoHide
    Clipboard := EN_Output
    TooltipTimed("Copied to clipboard.", 1200)
return

EN_Insert:
    Gui, EN:Submit, NoHide
    ReplaceSelection(EN_Output)
return

EN_Close:
    Gui, EN:Destroy
return

; =====================================================================================
; TEXT + CLIPBOARD HELPERS (PowerScribe / Epic integration assumptions)
; =====================================================================================
GetActiveReportText(){
    ; Attempts to grab all text from the active control while preserving the clipboard.
    SavedClip := ClipboardAll
    Clipboard :=

    ; Pass 1: try focused control text (fastest, avoids flashing selection)
    ControlGetFocus, ctrl, A
    if (ctrl != ""){
        ControlGetText, text, %ctrl%, A
        if (text != ""){
            Clipboard := SavedClip
            return text
        }
    }

    ; Pass 2: select-all/copy
    Send ^a
    Sleep, 200
    Send ^c
    ClipWait, 1
    text := Clipboard
    Clipboard := SavedClip

    if (text = ""){
        MsgBox, 48, Capture Error
        , % "Could not read text from the active window.`n" . "Click inside the report text box, ensure it is editable, then retry."
    }
    return text
}

GetSelectionOrReport(){
    sel := GetSelectedText()
    if (sel != "")
        return sel
    return GetActiveReportText()
}

GetSelectedText(){
    SavedClip := ClipboardAll
    Clipboard :=
    Send ^c
    ClipWait, 0.3
    text := Clipboard
    Clipboard := SavedClip
    return text
}

GetSelectedOrClipboardText(){
    txt := GetSelectedText()
    if (txt = "")
        txt := Clipboard
    return txt
}

ReplaceSelection(text, colorBGR:=""){
    ; Replaces the selection if present; otherwise inserts at caret.
    ; If a color is provided and we're in the PowerScribe text box, highlight the inserted text.
    global pscribetxtbox, pscribehwnd

    targetHwnd := 0, startSel := 0, endSel := 0
    if (colorBGR != ""){
        ; Attempt to capture the current selection bounds inside PowerScribe.
        if (pscribetxtbox = "" || !pscribehwnd)
            GetPowerscribeTextboxname()
        ControlGetFocus, focusedCtrl, PowerScribe 360 | Reporting
        if (pscribehwnd && focusedCtrl = pscribetxtbox){
            targetHwnd := pscribehwnd
            Edit_GetSel(targetHwnd, startSel, endSel)
            ControlFocus, %pscribetxtbox%, PowerScribe 360 | Reporting
        }
    }

    SavedClip := ClipboardAll
    Clipboard := text
    Send ^v
    Sleep, 80
    Clipboard := SavedClip

    if (targetHwnd){
        newEnd := startSel + StrLen(text)
        SetTextColorSafe(targetHwnd, startSel, newEnd, colorBGR)
        Edit_SetSel(targetHwnd, newEnd, newEnd)  ; return caret to end of the inserted text
    }
}

ReplaceEntireText(text){
    global pscribetxtbox, pscribehwnd
    if (pscribetxtbox = "" || !pscribehwnd)
        GetPowerscribeTextboxname()
    if (pscribetxtbox != ""){
        ControlSetText, %pscribetxtbox%, %text%, PowerScribe 360 | Reporting
        ControlFocus, %pscribetxtbox%, PowerScribe 360 | Reporting
        return
    }
    SavedClip := ClipboardAll
    Clipboard := text
    Send ^a
    Sleep, 60
    Send ^v
    Sleep, 80
    Clipboard := SavedClip
}

JumpToText(needle){
    ; Light-weight find: opens Ctrl+F, pastes needle, presses Enter.
    SavedClip := ClipboardAll
    Clipboard := needle
    Send ^f
    Sleep, 80
    Send ^v
    Sleep, 80
    Send {Enter}
    Sleep, 120
    Clipboard := SavedClip
}

AddHistory(label, text){
    global HistoryStack, HistoryLimit
    if (text = ""){
        text := GetPowerscribeReportText()
    }
    if (text = "")
        return
    HistoryStack.Push({"Label": label, "Text": text, "When": A_Now})
    while (HistoryStack.MaxIndex() > HistoryLimit)
        HistoryStack.RemoveAt(1)
}

; =====================================================================================
; GPT CALLER (provider-agnostic)
; =====================================================================================
EnsureApiConfig(){
    global Settings
    if (Settings.Provider != "Ollama" && Settings.ApiKey = ""){
        MsgBox, 48, API Settings, Enter an API key in Preferences before using GPT features.
        return false
    }
    if (Settings.Endpoint = "" || Settings.Model = ""){
        MsgBox, 48, API Settings, Endpoint and model are required.
        return false
    }
    return true
}

CallChatModel(systemPrompt, userText, temperature:=0.2, maxTokens:=800){
    global Settings
    try {
        messages := [ {"role":"system","content":systemPrompt}
                    , {"role":"user","content":userText} ]
        payload := { "messages": messages
                   , "temperature": temperature }
        if (Settings.Provider != "Ollama")
            payload["max_tokens"] := maxTokens
        payload["model"] := Settings.Model
        if (Settings.Provider = "Ollama"){
            payload["stream"] := false
        }

        jsonBody := Jxon_Dump(payload)
        http := ComObjCreate("WinHttp.WinHttpRequest.5.1")
        http.Open("POST", Settings.Endpoint, false)
        http.SetTimeouts(5000, 5000, 30000, 30000)
        http.SetRequestHeader("Content-Type","application/json")
        if (Settings.Provider = "OpenAI")
            http.SetRequestHeader("Authorization","Bearer " Settings.ApiKey)
        else if (Settings.Provider = "Azure OpenAI")
            http.SetRequestHeader("api-key", Settings.ApiKey)
        ; Ollama typically needs no key

        http.Send(jsonBody)
        http.WaitForResponse()
        status := http.Status
        resp := http.ResponseText
    } catch e {
        MsgBox, 16, API Error, % "Failed to call endpoint.`n" e.Message
        return ""
    }

    if (status < 200 || status >= 300){
        MsgBox, 16, API Error, % "HTTP " status "`n" resp
        return ""
    }

    result := ParseChatResponse(resp, Settings.Provider)
    if (result = ""){
        MsgBox, 16, API Error, % "Unable to parse response.`nRaw:`n" resp
        return ""
    }
    return result
}

ParseChatResponse(jsonText, provider){
    obj := Jxon_Load(jsonText)
    if (!IsObject(obj))
        return ""
    if (provider = "Ollama"){
        try return obj["message"]["content"]
    } else {
        try return obj["choices"][1]["message"]["content"]
    }
    return ""
}

; =====================================================================================
; REPORT UTILITIES
; =====================================================================================
BuildReportWithImpression(report, impression){
    ; Ensures "IMPRESSION:" exists once and replaces/creates content.
    cleanImpr := Trim(impression)
    if (SubStr(cleanImpr, 1, 11) != "IMPRESSION:")
        cleanImpr := "IMPRESSION:`r`n" cleanImpr

    pos := RegExMatch(report, "i)IMPRESSION:")
    if (!pos){
        return report "`r`n`r`n" cleanImpr
    }
    ; Replace anything after IMPRESSION:
    before := SubStr(report, 1, pos-1)
    return before cleanImpr
}

; =====================================================================================
; POWERSCRIBE HELPERS (minimal copies from legacy)
; =====================================================================================
GetPowerscribeTextboxname(){
    global pscribetxtbox, pscribetxtbox2, pscribehwnd
    pscribetxtbox := ""
    pscribetxtbox2 := ""
    bestLen := -1
    secondLen := -1
    WinGet, CtrlList, ControlList, PowerScribe 360 | Reporting
    Loop, Parse, CtrlList, `n
        if (InStr(A_LoopField, "RICHEDIT50W")){
            ControlGetText, t, %A_LoopField%, PowerScribe 360 | Reporting
            len := StrLen(t)
            if (len > bestLen){
                pscribetxtbox2 := pscribetxtbox
                secondLen := bestLen
                pscribetxtbox := A_LoopField
                bestLen := len
            } else if (len > secondLen){
                pscribetxtbox2 := A_LoopField
                secondLen := len
            }
        }
    
    if (pscribetxtbox != "")
        ControlGet, pscribehwnd, Hwnd,, %pscribetxtbox%, PowerScribe 360 | Reporting
    return pscribetxtbox
}

GetPowerscribeReportText(){
    global pscribetxtbox, pscribehwnd
    if (pscribetxtbox = "" || !pscribehwnd)
        GetPowerscribeTextboxname()
    if (pscribetxtbox = "" || !pscribehwnd)
        return ""
    SavedClip := ClipboardAll
    Clipboard :=

    ControlFocus, %pscribetxtbox%, PowerScribe 360 | Reporting
    Sleep, 100
    ControlSend,, ^a, ahk_id %pscribehwnd%
    Sleep, 150
    ControlSend,, ^c, ahk_id %pscribehwnd%
    ClipWait, 1
    text := Clipboard
    Clipboard := SavedClip
    return text
}

PowerscribeSendText2Target(TargetString, ReplaceString){
    global pscribetxtbox, pscribehwnd, Settings
    if (pscribetxtbox = "" || !pscribehwnd){
        GetPowerscribeTextboxname()
        if (pscribetxtbox = "" || !pscribehwnd){
            MsgBox, 48, PowerScribe, Could not locate the PowerScribe text box.
            return false
        }
    }
    ControlGetText, ReportText, %pscribetxtbox%, PowerScribe 360 | Reporting
    if (ReportText = "")
        ReportText := ""

    pos := RegExMatch(ReportText, "i)" . TargetString)
    if (pos){
        before := SubStr(ReportText, 1, pos-1)
        newText := before . TargetString . "`r`n" . ReplaceString
        ; Calculate cursor position at the start of the IMPRESSION section
        cursorPos := pos - 1
        ; Calculate position for coloring
        colorStart := pos - 1 + StrLen(TargetString) + 2
        colorEnd := colorStart + StrLen(ReplaceString)
    } else {
        newText := ReportText . "`r`n`r`n" . TargetString . "`r`n" . ReplaceString
        ; Calculate cursor position at the start of the newly added IMPRESSION section
        cursorPos := StrLen(ReportText) + 4
        ; Calculate position for coloring
        colorStart := StrLen(ReportText) + 4 + StrLen(TargetString) + 2
        colorEnd := colorStart + StrLen(ReplaceString)
    }

    ; Set the complete text
    ControlSetText, %pscribetxtbox%, %newText%, PowerScribe 360 | Reporting
    ControlFocus, %pscribetxtbox%, PowerScribe 360 | Reporting
    Sleep, 150

    ; Try to color the text (non-critical, don't crash if it fails)
    try {
        SetTextColorSafe(pscribehwnd, colorStart, colorEnd, Settings.AITextColor)
    }

    ; Move cursor to the IMPRESSION section and scroll it into view
    Sleep, 50
    Edit_SetSel(pscribehwnd, cursorPos, cursorPos)

    return true
}

; =====================================================================================
; IMPRESSION BUILDERS
; =====================================================================================
OnGenerateImpression_GPT(reportText){
    sysPrompt := "You are a radiology assistant. Create a concise IMPRESSION section using bullet points or numbered items when appropriate. Preserve measurements, laterality, dates, and comparison details. Do not invent data."
    return CallChatModel(sysPrompt, reportText, 0.2, 800)
}

GPT_CleanImpression(GPT_Impression_Raw, GPT_Impression_Numbered) {
    ; Remove "IMPRESSION:" header and normalize line breaks/whitespace
    text := StrReplace(GPT_Impression_Raw, "IMPRESSION:", "")
    text := StrReplace(text, "`r`n", "`n")
    text := StrReplace(text, "`r", "`n")
    text := RegExReplace(text, "^\s+")            ; trim top
    text := RegExReplace(text, "\s+$")            ; trim bottom
    text := RegExReplace(text, "`n{2,}", "`n")    ; collapse extra blank lines

    cleaned := ""
    idx := 0

    Loop, Parse, text, `n
    {
        line := A_LoopField

        ; 1) Strip bullets like '-', '*', '•' with trailing spaces
        line := RegExReplace(line, "^\s*(?:[-*•]\s+)+", "")

        ; 2) Strip one or more integer-based list tokens (e.g., "1. ", "2) ", "1. 1. ")
        ;    Pattern ensures a space after '.' or ')' so decimals like '0.9' remain untouched:
        ;    ^\s*(?:\d+\s*[.)]\s+)+
        line := RegExReplace(line, "^\s*(?:\d+\s*[.)]\s+)+", "")

        ; 3) Optional: trim again (after stripping markers)
        line := RegExReplace(line, "^\s+")
        line := RegExReplace(line, "\s+$")

        if (line = "")
        {
            ; keep blank lines unnumbered
            cleaned .= (cleaned ? "`r`n" : "") . ""
            continue
        }

        if (GPT_Impression_Numbered) {
            idx++
            line := idx ". " line
        }

        cleaned .= (cleaned ? "`r`n" : "") . line
    }

    ; Return cleaned text WITHOUT "IMPRESSION:" header
    ; PowerscribeSendText2Target will add it
    return cleaned
}

; =====================================================================================
; SETTINGS / PERSISTENCE
; =====================================================================================
EnsureConfigDir(){
    global ConfigDir
    if (!FileExist(ConfigDir))
        FileCreateDir, % ConfigDir
}

LoadSettings(){
    global Settings, DefaultSettings, ConfigFile, DefaultEndpoints
    Settings := DefaultSettings.Clone()
    if (!FileExist(ConfigFile)){
        SaveSettings()
        return
    }
    for k, v in Settings {
        IniRead, val, % ConfigFile, API, %k%, %v%
        Settings[k] := val
    }
    if (Settings.Endpoint = "")
        Settings.Endpoint := DefaultEndpoints[Settings.Provider]
}

SaveSettings(){
    global Settings, ConfigFile
    for k, v in Settings
        IniWrite, %v%, % ConfigFile, API, %k%
}

LoadPrompts(){
    global Prompts, ConfigFile, DefaultPrompts
    Prompts := []
    IniRead, count, % ConfigFile, Prompts, Count, 0
    if (count = 0){
        Prompts := DefaultPrompts.Clone()
        SavePrompts()
        return
    }
    Loop, %count% {
        IniRead, name, % ConfigFile, Prompts, Prompt%A_Index%_Name, 
        IniRead, text, % ConfigFile, Prompts, Prompt%A_Index%_Text, 
        IniRead, hk, % ConfigFile, Prompts, Prompt%A_Index%_Hotkey, 
        if (name != "" && text != "")
            Prompts.Push({"Name": name, "Text": text, "Hotkey": hk})
    }
    RegisterPromptHotkeys()
}

SavePrompts(){
    global Prompts, ConfigFile
    IniWrite, % Prompts.MaxIndex(), % ConfigFile, Prompts, Count
    Loop, % Prompts.MaxIndex(){
        IniWrite, % Prompts[A_Index].Name, % ConfigFile, Prompts, Prompt%A_Index%_Name
        IniWrite, % Prompts[A_Index].Text, % ConfigFile, Prompts, Prompt%A_Index%_Text
        IniWrite, % Prompts[A_Index].Hotkey, % ConfigFile, Prompts, Prompt%A_Index%_Hotkey
    }
    RebuildCustomPromptMenu()
    RegisterPromptHotkeys()
}

LoadHotkeys(){
    global Hotkeys, DefaultHotkeys, ConfigFile
    Hotkeys := DefaultHotkeys.Clone()
    if (!FileExist(ConfigFile)){
        SaveHotkeys()
        return
    }
    for k, v in Hotkeys {
        IniRead, val, % ConfigFile, Hotkeys, %k%, %v%
        Hotkeys[k] := val
    }
}

SaveHotkeys(){
    global Hotkeys, ConfigFile
    for k, v in Hotkeys
        IniWrite, %v%, % ConfigFile, Hotkeys, %k%
}

UpdateStatusLine(){
    global Settings
    GuiControl, Main:, StatusLine, % "Provider: " Settings.Provider " | Model: " Settings.Model " | Endpoint: " Settings.Endpoint
}

; =====================================================================================
; STARTUP DISCLAIMER
; =====================================================================================
ShowStartupEula(){
    global SessionEulaShown, AppName
    if (SessionEulaShown)
        return
    text =
    (
    AI output is for educational demonstration only.
    No warranties are provided. Always verify content before signing.
    AI may produce errors, omissions, or hallucinations.
    Handle PHI in compliance with HIPAA and local policies.
    )
    MsgBox, 64, %AppName% Notice, %text%
    SessionEulaShown := true
}

; =====================================================================================
; SMALL UTILITIES
; =====================================================================================
TooltipTimed(msg, duration:=1000){
    Tooltip, %msg%
    SetTimer, __HideTip, -%duration%
}
__HideTip:
    Tooltip
return

; Normalize any color value to a 6-digit hex string (no 0x prefix)
NormalizeColor6(colorVal) {
    if (colorVal = "")
        return "000000"
    if (colorVal is integer)
        return Format("{:06X}", colorVal & 0xFFFFFF)
    colorVal := RegExReplace(colorVal, "^0x", "")
    colorVal := RegExReplace(colorVal, "[^0-9A-Fa-f]")
    return SubStr("000000" colorVal, -6)
}

; Color conversion utilities (Windows uses BGR, AHK Gui uses RGB)
BGR2RGB(bgr) {
    ; Convert BGR (0xBBGGRR) to RGB (0xRRGGBB)
    bgr := NormalizeColor6(bgr)
    b := SubStr(bgr, 1, 2)
    g := SubStr(bgr, 3, 2)
    r := SubStr(bgr, 5, 2)
    return r . g . b
}

RGB2BGR(rgb) {
    ; Convert RGB (0xRRGGBB) to BGR (0xBBGGRR)
    rgb := NormalizeColor6(rgb)
    r := SubStr(rgb, 1, 2)
    g := SubStr(rgb, 3, 2)
    b := SubStr(rgb, 5, 2)
    return "0x" . b . g . r
}

ChooseColor(defaultRGB := "FF8000") {
    ; Show Windows color picker dialog
    ; Returns RGB hex string (e.g., "FF8000") or empty string if canceled
    defaultRGB := NormalizeColor6(defaultRGB)

    ; Convert RGB hex string to BGR decimal for Windows
    r := "0x" . SubStr(defaultRGB, 1, 2)
    g := "0x" . SubStr(defaultRGB, 3, 2)
    b := "0x" . SubStr(defaultRGB, 5, 2)
    defaultBGR := (b << 16) | (g << 8) | r

    ; Allocate custom colors array (prefill with the current color)
    VarSetCapacity(custColors, 64, 0)
    Loop, 16
        NumPut(defaultBGR, custColors, (A_Index-1)*4, "UInt")

    ; Build CHOOSECOLOR structure with proper padding for 32/64-bit
    if (A_PtrSize = 8) {
        VarSetCapacity(CHOOSECOLOR, 72, 0)               ; sizeof(CHOOSECOLOR) on 64-bit
        NumPut(72, CHOOSECOLOR, 0, "UInt")               ; lStructSize
        NumPut(0, CHOOSECOLOR, 8, "Ptr")                 ; hwndOwner (none)
        NumPut(0, CHOOSECOLOR, 16, "Ptr")                ; hInstance
        NumPut(defaultBGR, CHOOSECOLOR, 24, "UInt")      ; rgbResult (init color)
        NumPut(&custColors, CHOOSECOLOR, 32, "Ptr")      ; lpCustColors
        NumPut(0x00000103, CHOOSECOLOR, 40, "UInt")      ; Flags: CC_RGBINIT | CC_FULLOPEN | CC_ANYCOLOR
        selectedOffset := 24
    } else {
        VarSetCapacity(CHOOSECOLOR, 36, 0)               ; sizeof(CHOOSECOLOR) on 32-bit
        NumPut(36, CHOOSECOLOR, 0, "UInt")               ; lStructSize
        NumPut(0, CHOOSECOLOR, 4, "Ptr")                 ; hwndOwner (none)
        NumPut(0, CHOOSECOLOR, 8, "Ptr")                 ; hInstance
        NumPut(defaultBGR, CHOOSECOLOR, 12, "UInt")      ; rgbResult (init color)
        NumPut(&custColors, CHOOSECOLOR, 16, "Ptr")      ; lpCustColors
        NumPut(0x00000103, CHOOSECOLOR, 20, "UInt")      ; Flags: CC_RGBINIT | CC_FULLOPEN | CC_ANYCOLOR
        selectedOffset := 12
    }

    fn := "comdlg32\ChooseColor" (A_IsUnicode ? "W" : "A")
    result := DllCall(fn, "Ptr", &CHOOSECOLOR)
    if (result) {
        selectedBGR := NumGet(CHOOSECOLOR, selectedOffset, "UInt")
        ; Convert BGR to RGB hex
        r := Format("{:02X}", (selectedBGR & 0xFF))
        g := Format("{:02X}", ((selectedBGR >> 8) & 0xFF))
        b := Format("{:02X}", ((selectedBGR >> 16) & 0xFF))
        return r . g . b
    }
    return ""
}

; Allow dragging of windows with no title bar
WM_LBUTTONDOWN() {
    global ErrCtrlMap
    ; Get the control that was clicked
    MouseGetPos,,, winHwnd, controlHwnd, 2

    ; If it's a clickable error control, don't drag - let the g-label fire
    if (ErrCtrlMap.HasKey(controlHwnd))
        return

    ; Otherwise allow dragging
    PostMessage, 0xA1, 2,,, A
}

_:  ; placeholder for disabled menu items
return

ExitScript:
    ExitApp
return

; =====================================================================================
; MINIMAL JSON (JXON) HELPERS
; -------------------------------------------------------------------------------------
; Compact JSON encode/decode helpers adapted for single-file portability.
; Supports objects, arrays, numbers, strings, true/false/null.
; =====================================================================================
Jxon_Load(ByRef src, rev=false){
    pos := 1
    return Jxon_ParseValue(src, pos)
}

Jxon_ParseValue(ByRef src, ByRef pos){
    Jxon_SkipWS(src, pos)
    ch := SubStr(src, pos, 1)
    if (ch = """")
        return Jxon_ParseString(src, pos)
    else if (ch = "{")
        return Jxon_ParseObject(src, pos)
    else if (ch = "[")
        return Jxon_ParseArray(src, pos)
    else if (SubStr(src, pos, 4) = "true")
        pos += 4, Jxon_SkipWS(src, pos), return true
    else if (SubStr(src, pos, 5) = "false")
        pos += 5, Jxon_SkipWS(src, pos), return false
    else if (SubStr(src, pos, 4) = "null")
        pos += 4, Jxon_SkipWS(src, pos), return ""
    else {
        RegExMatch(SubStr(src, pos), "^-?\d+(\.\d+)?([eE][+-]?\d+)?", m)
        pos += StrLen(m)
        Jxon_SkipWS(src, pos)
        return m
    }
}

Jxon_ParseString(ByRef src, ByRef pos){
    pos++ ; skip opening quote
    out := ""
    while (pos <= StrLen(src)){
        ch := SubStr(src, pos, 1)
        if (ch = """"){
            pos++
            Jxon_SkipWS(src, pos)
            return out
        } else if (ch = "\"){
            pos++
            esc := SubStr(src, pos, 1)
            if (esc="n")
                out.="`n"
            else if (esc="r")
                out.="`r"
            else if (esc="t")
                out.="`t"
            else if (esc="""")
                out.=""""
            else if (esc="\")
                out.="\"
            else
                out.=esc
            pos++
        } else {
            out .= ch
            pos++
        }
    }
    return out
}

Jxon_ParseArray(ByRef src, ByRef pos){
    arr := []
    pos++ ; skip [
    Jxon_SkipWS(src, pos)
    if (SubStr(src, pos, 1) = "]"){
        pos++
        Jxon_SkipWS(src, pos)
        return arr
    }
    loop {
        val := Jxon_ParseValue(src, pos)
        arr.Push(val)
        Jxon_SkipWS(src, pos)
        ch := SubStr(src, pos, 1)
        if (ch = ","){
            pos++
            continue
        } else if (ch = "]"){
            pos++
            Jxon_SkipWS(src, pos)
            break
        } else break
    }
    return arr
}

Jxon_ParseObject(ByRef src, ByRef pos){
    obj := {}
    pos++ ; skip {
    Jxon_SkipWS(src, pos)
    if (SubStr(src, pos, 1) = "}"){
        pos++
        Jxon_SkipWS(src, pos)
        return obj
    }
    loop {
        key := Jxon_ParseString(src, pos)
        if (SubStr(src, pos, 1) != ":")
            break
        pos++
        val := Jxon_ParseValue(src, pos)
        obj[key] := val
        Jxon_SkipWS(src, pos)
        ch := SubStr(src, pos, 1)
        if (ch = ","){
            pos++
            continue
        } else if (ch = "}"){
            pos++
            Jxon_SkipWS(src, pos)
            break
        } else break
    }
    return obj
}

Jxon_SkipWS(ByRef src, ByRef pos){
    while (pos <= StrLen(src)){
        ch := SubStr(src, pos, 1)
        if (ch != " " && ch != "`t" && ch != "`r" && ch != "`n")
            break
        pos++
    }
}

Jxon_Dump(obj){
    if (!IsObject(obj)){
        if (obj = "")
            return "null"
        if obj is number
            return obj
        return """" . Jxon_Escape(obj) . """"
    }
    isArray := (obj.MaxIndex() != "")
    if (isArray){
        out := "["
        for each, val in obj
            out .= Jxon_Dump(val) ","
        return RTrim(out, ",") "]"
    } else {
        out := "{"
        for k, v in obj
            out .= """" Jxon_Escape(k) """:" . Jxon_Dump(v) ","
        return RTrim(out, ",") "}"
    }
}

Jxon_Escape(str){
    str := StrReplace(str, "\", "\\")
    str := StrReplace(str, """", "\""")
    str := StrReplace(str, "`r", "\r")
    str := StrReplace(str, "`n", "\n")
    str := StrReplace(str, "`t", "\t")
    return str
}

GetPowerscribeElementInfo(Element){
	OpenPowerscribeSidePanel()
	;GetPowerscribeControlbyText only returns the control name of the actual Element itself but not the corresponding value in the adjacent control
	;this function will return something like GE12345679 if Element is "MRN:"
	tempstr := StrSplit(GetPowerscribeControlbyText(Element), "_ad")
	;get the control name for the adjacent control
	nextcontrolnumber := tempstr[2] - 1
	;reassemable the control name
	nextcontrol := tempstr[1] . "_ad" . nextcontrolnumber
	Controlgettext, ElementInfo, %nextcontrol%, PowerScribe 360 | Reporting
	Return ElementInfo
}

OpenPowerscribeSidePanel(){
	WinGet, MMX2, MinMax, PowerScribe 360 | Reporting
	if (MMX2 == -1)
		WinRestore, PowerScribe 360 | Reporting
	;WinActivate, PowerScribe 360 | Reporting
	WinGetText, panel, PowerScribe 360 | Reporting
	if (!InStr(panel, ".)) - ")) {
			ControlSend, Main Menu, {Alt Down}{v}{o}, PowerScribe 360 | Reporting
			Sleep 50
			ControlSend, Main Menu, {Alt Up}, PowerScribe 360 | Reporting
	}
	WinGetText, panel, PowerScribe 360 | Reporting
	isSuccess := InStr(panel, ".)) - ") || InStr(panel, "TEMPORARY") 
	return isSuccess
}

GetPowerscribeControlbyText(Text){
	WinGet, PowerscribeControls, ControlList, PowerScribe 360 | Reporting
	Loop, Parse, PowerscribeControls, `n
	{
		ControlGetText, TempControlText, %A_LoopField%, PowerScribe 360 | Reporting
		If(TempControlText = Text)
		{
			ControlName := A_LoopField
			Break
		}
	}
	Return ControlName
}
