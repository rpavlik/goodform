
-- Uses LuaXml http://viremo.eludi.net/LuaXML/ and LuaZip http://www.keplerproject.org/luazip/manual.html

-- Some WordprocessingML docs that might be useful: search for "complex fields" here:
-- http://rep.oio.dk/microsoft.com/officeschemas/wordprocessingml_article.htm

require "LuaXml"




local function findWithPredicate(list, first, last, predicate)
	for childnum=first, last do
		if predicate(list[childnum]) then
			return childnum, list[childnum]
		end
	end
end

local function fldCharTypeChecker(fldCharType)
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

local isRun = function(elem)
	return type(elem) == "table" and elem:tag() == "w:r"
end

local FieldTypes = {
	["w:textInput"] = {
		init = function(self)
			self.separateID = findWithPredicate(self.block, self.startID, self.endID, isFieldSeparate)
			self.contentID = findWithPredicate(self.block, self.separateID, self.endID, isRun)
			return self
		end,
		__tostring = function(self)
			return "Text Input: name='" .. self.name .. "'"
		end
	},
	["w:checkBox"] = {
		init = function(self)
			return self
		end,
		__tostring = function(self)
			return "Checkbox: name='" .. self.name .. "'"
		end
	},
	["w:ddList"] = {
		init = function(self, elem)
			self.entries = {}
			for _, v in ipairs(elem) do
				table.insert(self.entries, v["w:val"])
			end
			return self
		end,
		__tostring = function(self)
			return "Drop-down list: name='" .. self.name .. "', entries: " .. table.concat(self.entries, ", ")
		end
	}
}

local function Field(block, startID)
	assert(isFieldStart(block[startID]), "Field must be initialized with a block and a fldchar begin")
	local self = {
		["block"] = block,
		["startID"] = startID,
		["endID"] = findWithPredicate(block, startID, #block, isFieldEnd),
		["name"] = block[startID]
			:find("w:fldChar", "w:fldCharType", "begin")
			:find("w:name")
			["w:val"]
	}
	local ffData = block[startID]:find("w:ffData")
	for _, elem in ipairs(ffData) do
		if FieldTypes[elem:tag()] ~= nil then
			setmetatable(self, FieldTypes[elem:tag()])
			return FieldTypes[elem:tag()].init(self, elem)
		end
	end
	print "Got a form field, but couldn't determine the type!"
	return nil
end


--- Temporary place to put all stuff.
local allfields = {}

function recursivelyFindFields(elem)
	--print(elem:tag())
	local fieldname
	for childnum, child in ipairs(elem) do
		if isFieldStart(child) then
			--print("Found start of field", fieldname)
			table.insert(allfields, Field(elem, childnum))
		elseif type(child) == "table" then
			recursivelyFindFields(child)
		end
	end
end

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

recursivelyFindFields(xfile)

for i, field in ipairs(allfields) do
	print(i, field)
end

