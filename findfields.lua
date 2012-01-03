
require "LuaXml"




function isFieldStart(elem)
	if type(elem) == "table" and elem:tag() == "w:r" then
		local fldChar = elem:find("w:fldChar", "w:fldCharType", "begin")
		if fldChar ~= nil then
			local name = fldChar:find("w:name")
			if name ~= nil then
				return name["w:val"]
			end
		end
	end
end

function handleField(block, runID)
	local name = block[runID]
		:find("w:fldChar", "w:fldCharType", "begin")
		:find("w:name")
		["w:val"]
	local separateRunID
	local contentRunID
	local endRunID
	for childnum=runID, #block do
		if block[childnum]:tag() == "w:r" then
			if separateRunID == nil
				and block[childnum]:find("w:fldChar", "w:fldCharType", "separate") then
					separateRunID = childnum
			elseif contentRunID == nil then
				contentRunID = childnum
			elseif block[childnum]:find("w:fldChar", "w:fldCharType", "end") then
				endRunID = childnum
			end
		end
	end
	if separateRunID == nil then
		print("Got field", name, "without a separate fldchartype.", block)
		return
	end
	contentRunID = separateRunID + 1
	local contentRun = block[contentRunID]
	print("Got field", name, "with content", contentRun)

end

function recursivelyFindFields(elem)
	--print(elem:tag())
	local fieldname
	for childnum, child in ipairs(elem) do
		fieldname = isFieldStart(child)
		if fieldname ~= nil then
			--print("Found start of field", fieldname)
			handleField(elem, childnum)
		elseif type(child) == "table" then
			recursivelyFindFields(child)
		end
	end
end
--[[
for k,v in pairs(xfile) do
	print(k)
end
print(xfile:tag())]]
--print(xfile)

require "zip"
local zfile, err = zip.open(arg[1] or "minimalform.docx")
if not zfile then
	error("Could not open zip file: " .. err)
end

local zdocfile, err = zfile:open("word/document.xml")
if not zdocfile then
	error("Could not open word/document.xml in zip file: " .. err)
end

local xfile = xml.eval(zdocfile:read("*a"))
zdocfile:close()
zfile:close()
local body = xfile[1]

recursivelyFindFields(xfile)
