findfields = require "findfields"


xfile = findfields.loadDOCX(arg[1] or "minimalform.docx")

for field in findfields.iterateFields(xfile) do
	print(field)
end
