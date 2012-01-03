
require "LuaXml"



function fldCharTypeChecker(fldCharType)
	return function (elem)
		if type(elem) == "table" and elem:tag() == "w:r" then
			local fldChar = elem:find("w:fldChar", "w:fldCharType", fldCharType)
			return fldChar ~= nil
		end
		return false
	end
end

local isFieldStart = fldCharTypeChecker("begin")
local isFieldSeparate = fldCharTypeChecker("separate")
local isFieldEnd = fldCharTypeChecker("end")

function handleField(block, runID)
	local name = block[runID]
		:find("w:fldChar", "w:fldCharType", "begin")
		:find("w:name")
		["w:val"]
	local separateRunID
	local contentRunID
	local endRunID
	print("Starting to look at " .. tostring(runID))
	print(block[runID])
	for childnum=runID, #block do
		print("Child " .. tostring(childnum))
		if separateRunID == nil then
			if isFieldSeparate(block[childnum]) then
				print("Got separate")
				separateRunID = childnum
			end
		elseif contentRunID == nil then
			if type(block[childnum]) == "table" and block[childnum]:tag() == "w:r" then
				print("Got content")
				contentRunID = childnum
			end
		elseif endRunID == nil then
			if isFieldEnd(block[childnum]) then

				print("Got end")
				endRunID = childnum
			end
		end
	end

	if separateRunID == nil then
		print("Got field", name, "without a separate fldchartype.", block)
		return
	end

	local contentRun = block[contentRunID]
	print("Got field", name, "with fldChars", runID, separateRunID, contentRunID, endRunID, " and content", contentRun)

end

function recursivelyFindFields(elem)
	--print(elem:tag())
	local fieldname
	for childnum, child in ipairs(elem) do
		if isFieldStart(child) then
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
