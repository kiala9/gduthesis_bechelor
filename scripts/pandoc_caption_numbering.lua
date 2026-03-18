local chapter_no = 0
local figure_no = 0
local table_no = 0
local in_appendix = false
local appendix_label = nil

local function stringify_inlines(inlines)
  return pandoc.utils.stringify(inlines or {})
end

local function make_prefix(kind)
  local chapter_label = appendix_label or tostring(chapter_no)
  if kind == "fig" then
    figure_no = figure_no + 1
    return "图" .. chapter_label .. "." .. tostring(figure_no) .. " "
  end
  table_no = table_no + 1
  return "表" .. chapter_label .. "." .. tostring(table_no) .. " "
end

function Header(h)
  if h.level == 1 then
    local text = pandoc.utils.stringify(h.content)
    local appendix_match = text:match("^附录([A-Z])")
    if appendix_match then
      in_appendix = true
      appendix_label = appendix_match
      figure_no = 0
      table_no = 0
      return h
    end
    if not h.classes:includes("unnumbered") then
      chapter_no = chapter_no + 1
      figure_no = 0
      table_no = 0
      if not in_appendix then
        appendix_label = nil
      end
    end
  end
  return h
end

function Figure(fig)
  local blocks = fig.caption.long
  if #blocks == 0 then
    return fig
  end
  local first = blocks[1]
  if first.t == "Plain" or first.t == "Para" then
    local caption_text = stringify_inlines(first.content)
    if not caption_text:match("^图[%dA-Z]+%.%d+") then
      table.insert(first.content, 1, pandoc.Str(make_prefix("fig")))
    end
    blocks[1] = first
    fig.caption.long = blocks
  end
  return fig
end

function Table(tbl)
  local blocks = tbl.caption.long
  if #blocks == 0 then
    return tbl
  end
  local first = blocks[1]
  if first.t == "Plain" or first.t == "Para" then
    local caption_text = stringify_inlines(first.content)
    if not caption_text:match("^表[%dA-Z]+%.%d+") then
      table.insert(first.content, 1, pandoc.Str(make_prefix("table")))
    end
    blocks[1] = first
    tbl.caption.long = blocks
  end
  return tbl
end
