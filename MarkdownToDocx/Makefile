FILE=Headers.docx

default:
	osascript -e 'tell application "Microsoft Word" to close (documents whose name is "${FILE}") without saving'
	dotnet run
	open ${FILE}