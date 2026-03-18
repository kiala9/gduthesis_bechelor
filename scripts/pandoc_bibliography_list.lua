local function is_bibliography_header(block)
  if block.t ~= "Header" then
    return false
  end
  local text = pandoc.utils.stringify(block)
  return text == "参考文献" or text == "References"
end

local function trim_manual_label(inlines)
  local text = pandoc.utils.stringify(inlines)
  local stripped = text:gsub("^%s*%[%d+%]%s*", "", 1)
  if stripped == text then
    return inlines
  end
  return { pandoc.Str(stripped) }
end

function Pandoc(doc)
  local blocks = {}
  local i = 1

  while i <= #doc.blocks do
    local block = doc.blocks[i]
    table.insert(blocks, block)

    if is_bibliography_header(block) then
      local items = {}
      local j = i + 1

      while j <= #doc.blocks do
        local next_block = doc.blocks[j]
        if next_block.t == "Header" then
          break
        end
        if next_block.t == "Para" or next_block.t == "Plain" then
          table.insert(items, trim_manual_label(next_block.content))
        elseif next_block.t == "BulletList" or next_block.t == "OrderedList" then
          break
        end
        j = j + 1
      end

      if #items > 0 then
        for idx, content in ipairs(items) do
          local para = pandoc.Para(
            {
              pandoc.Str("[" .. tostring(idx) .. "]"),
              pandoc.Space(),
              table.unpack(content),
            },
            pandoc.Attr("", {}, { { "custom-style", "Bibliography" } })
          )
          table.insert(blocks, para)
        end
        i = j - 1
      end
    end

    i = i + 1
  end

  return pandoc.Pandoc(blocks, doc.meta)
end
