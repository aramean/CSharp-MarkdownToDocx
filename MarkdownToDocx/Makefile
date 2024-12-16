FILE_OUTPUT=Headers.docx

default:
	osascript -e 'tell application "Microsoft Word" to close (documents whose name is "${FILE_OUTPUT}") without saving'
	dotnet run input.md
	open ${FILE_OUTPUT}