local keyword_label = utf8.char(20851, 38190, 35789, 65306)
local en_keyword_patterns = { "^Keywords:%s*", "^Key words:%s*" }

local function starts_with_keyword(text)
  if text:match("^" .. keyword_label) or text:match("^关键词:") then
    return "cn"
  end
  for _, pattern in ipairs(en_keyword_patterns) do
    if text:match(pattern) then
      return "en"
    end
  end
  return nil
end

function Para(para)
  local text = pandoc.utils.stringify(para)
  local kind = starts_with_keyword(text)
  if not kind then
    return nil
  end

  if kind == "cn" then
    local suffix = text:gsub("^" .. keyword_label, "", 1):gsub("^关键词:", "", 1)
    return pandoc.Para({
      pandoc.Span({ pandoc.Str(keyword_label) }, { ["custom-style"] = "AbstractKeywordLabel" }),
      pandoc.Str(suffix),
    })
  end

  local label = text:match("^Keywords:") and "Keywords:" or "Key words:"
  local suffix = text
  for _, pattern in ipairs(en_keyword_patterns) do
    suffix = suffix:gsub(pattern, "", 1)
  end
  return pandoc.Para({
    pandoc.Span({ pandoc.Str(label) }, { ["custom-style"] = "EnglishAbstractKeywordLabel" }),
    pandoc.Space(),
    pandoc.Str(suffix),
  })
end
