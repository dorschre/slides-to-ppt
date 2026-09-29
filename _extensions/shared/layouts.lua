-- Slide layout handling for the beamer formats.
--
-- Reads the heading attributes `layout` and `subtitle` of every slide
-- (level 2 heading) and rewrites the slide content accordingly. The PPTX
-- output does not use this filter; tools/md2pptx.py reads the same attributes.
--
-- Brand "iis": the slides follow the PowerPoint template. Lengths below are
-- those of the PowerPoint page in cm and are scaled to the beamer page with
-- SCALE; the rules for columns, pictures and tables are the ones of md2pptx.

local IMAGE_COLUMN = 38 -- percent of the text width, image-left / image-right
local TEXT_COLUMN = 60
local COLUMN_GAP = 2

local SCALE = 16 / 33.867 -- PowerPoint page width -> beamer page width
local IIS = {
  text_width = 31.2,
  body_top = 4.73,
  content_bottom = 16.4,
  title_top = 1.10,
  gap = 0.4,
  column_gap = 1.07,
  half_width = 15.1, -- column of the two column layouts
  picture_left = 17.43, -- picture placeholder of layout 17
  picture_bottom = 17.09,
  page_width = 33.867,
  text_size = 14,
  code_size = 12,
  table_size = 12,
}

local function raw(tex)
  return pandoc.RawBlock("latex", tex)
end

local function cm(value)
  return string.format("%.3fcm", value * SCALE)
end

local function inlines_to_latex(inlines)
  local tex = pandoc.write(pandoc.Pandoc({ pandoc.Plain(inlines) }), "latex")
  return (tex:gsub("%s+$", ""))
end

local function blocks_to_latex(blocks)
  local tex = pandoc.write(pandoc.Pandoc(blocks), "latex")
  return (tex:gsub("%s+$", ""))
end

local function to_latex(text)
  local doc = pandoc.read(text, "markdown")
  return inlines_to_latex(pandoc.utils.blocks_to_inlines(doc.blocks))
end

local function is_notes(block)
  return block.t == "Div" and block.classes:includes("notes")
end

local function is_columns(block)
  return block.t == "Div" and block.classes:includes("columns")
end

-- Quarto turns a picture with caption into a div with raw LaTeX:
-- \begin{figure} ... image ... \caption{...} ... \end{figure}
local function is_figure_div(block)
  if block.t ~= "Div" then
    return false
  end
  local first = block.content[1]
  return first ~= nil and first.t == "RawBlock" and first.text:find("\\begin{figure}", 1, true) ~= nil
end

local function image_of(block)
  if block.t == "Figure" or is_figure_div(block) then
    local found
    block:walk({
      Image = function(img)
        found = found or img
      end,
    })
    return found
  end
  if (block.t == "Para" or block.t == "Plain") and #block.content == 1
    and block.content[1].t == "Image" then
    return block.content[1]
  end
  return nil
end

local function figure_div_caption(block)
  local caption, inside = pandoc.List(), false
  for _, child in ipairs(block.content) do
    if child.t == "Para" or child.t == "Plain" then
      for _, inline in ipairs(child.content) do
        if inline.t == "RawInline" then
          if inline.text:find("\\caption", 1, true) then
            inside = true
          elseif inline.text:find("}", 1, true) then
            inside = false
          end
        elseif inside and inline.t ~= "Image" then
          caption:insert(inline)
        end
      end
    end
  end
  return #caption > 0 and caption or nil
end

local function caption_of(block)
  if is_figure_div(block) then
    return figure_div_caption(block)
  end
  if block.t == "Figure" and block.caption and block.caption.long
    and #block.caption.long > 0 then
    return pandoc.utils.blocks_to_inlines(block.caption.long)
  end
  return nil
end

local function column(width, blocks)
  return pandoc.Div(blocks, pandoc.Attr("", { "column" }, { width = width .. "%" }))
end

local function columns(cols, classes)
  return pandoc.Div(cols, pandoc.Attr("", classes or { "columns" }))
end

local function column_divs(div)
  local cols = {}
  for _, block in ipairs(div.content) do
    if block.t == "Div" and block.classes:includes("column") then
      cols[#cols + 1] = block
    end
  end
  return cols
end

-- Splits the slide into its first image, the remaining content and the notes.
local function split_image(body)
  local image, rest, notes = nil, pandoc.List(), pandoc.List()
  for _, block in ipairs(body) do
    if is_notes(block) then
      notes:insert(block)
    elseif not image and image_of(block) then
      image = block
    else
      rest:insert(block)
    end
  end
  return image, rest, notes
end

-- --------------------------------------------------------------------------
-- FAU
-- --------------------------------------------------------------------------

-- Equal widths for 1 to 5 (or more) columns unless the author set a width.
local function spread_columns(div)
  if not is_columns(div) then
    return nil
  end
  local cols = column_divs(div)
  if #cols == 0 then
    return nil
  end
  local width = math.floor((100 - COLUMN_GAP * (#cols - 1)) / #cols)
  for _, col in ipairs(cols) do
    if not col.attributes.width then
      col.attributes.width = width .. "%"
    end
  end
  return div
end

local function image_beside_text(body, image_left)
  local image, rest, notes = split_image(body)
  if not image then
    io.stderr:write("layouts: slide with image layout has no image\n")
    return body
  end
  local img = column(IMAGE_COLUMN, { image })
  local txt = column(TEXT_COLUMN, rest)
  local result = pandoc.List({ columns(image_left and { img, txt } or { txt, img }) })
  result:extend(notes)
  return result
end

local function full_image(body)
  local image, rest, notes = split_image(body)
  if not image then
    io.stderr:write("layouts: slide with layout full-image has no image\n")
    return body
  end
  local src = image_of(image).src
  local result = pandoc.List({
    raw("\\begin{center}\\includegraphics[width=\\textwidth,height=0.78\\textheight,keepaspectratio]{"
      .. src .. "}\\end{center}"),
  })
  result:extend(rest)
  result:extend(notes)
  return result
end

local function fau_slide(header, body, layout)
  local options = { "t" }
  if layout == "blank" then
    options[#options + 1] = "blank"
    header.content = {}
    header.attributes.subtitle = nil
  elseif layout == "image-left" then
    body = image_beside_text(body, true)
  elseif layout == "image-right" then
    body = image_beside_text(body, false)
  elseif layout == "full-image" then
    body = full_image(body)
  elseif layout == "content-large" then
    -- the theme pins the list font size, so it has to be raised as well
    body:insert(1, raw("\\large"
      .. "\\setbeamerfont{itemize/enumerate body}{size=\\large}"
      .. "\\setbeamerfont{itemize/enumerate subbody}{size=\\normalsize}"
      .. "\\setbeamerfont{itemize/enumerate subsubbody}{size=\\normalsize}"))
  end
  -- walk() visits only descendants, so the blocks are wrapped first
  body = pandoc.Div(body):walk({ Div = spread_columns }).content
  return body, options
end

-- --------------------------------------------------------------------------
-- Fraunhofer IIS: same rules as tools/md2pptx.py
-- --------------------------------------------------------------------------

local function text_length(block)
  return utf8.len(pandoc.utils.stringify(block)) or #pandoc.utils.stringify(block)
end

-- as text_height() of md2pptx: 19 pt follow text paragraphs, nothing list items
local function paragraph_height(chars, width, size, indent, after)
  local per_line = math.max(8, math.floor((width / 2.54 * 72 - indent) / (size * 0.5)))
  local lines = math.max(1, math.ceil(chars / per_line))
  return (lines * size * 1.32 + after) / 72 * 2.54
end

local function all_strong(block)
  local seen = false
  for _, inline in ipairs(block.content) do
    if inline.t == "Strong" then
      seen = true
    elseif inline.t ~= "Space" and inline.t ~= "SoftBreak" then
      return false
    end
  end
  return seen
end

local list_height

-- Height in cm (PowerPoint page), nil for elements that adapt to the space left.
local function height_of(block, width)
  if image_of(block) then
    return nil
  end
  local size = IIS.text_size
  if block.t == "Para" or block.t == "Plain" then
    return paragraph_height(text_length(block), width, size, 0, all_strong(block) and 0 or 19)
  elseif block.t == "Header" then
    return (16 * 1.32 + 19) / 72 * 2.54
  elseif block.t == "BulletList" or block.t == "OrderedList" then
    return list_height(block, width, 1)
  elseif block.t == "CodeBlock" then
    local _, count = block.text:gsub("\n", "\n")
    return ((count + 1) * IIS.code_size * 1.2 + 16) / 72 * 2.54
  elseif block.t == "Table" then
    local simple = pandoc.utils.to_simple_table(block)
    local rows = #simple.rows + (#simple.headers > 0 and 1 or 0)
    return rows * (IIS.table_size * 1.2 + 10) / 72 * 2.54
  elseif is_columns(block) then
    local cols = column_divs(block)
    local inner = (width - IIS.column_gap * (#cols - 1)) / math.max(#cols, 1)
    local highest = 0
    for _, col in ipairs(cols) do
      local total = 0
      for _, child in ipairs(col.content) do
        local h = height_of(child, inner)
        if h == nil then
          return nil
        end
        total = total + h
      end
      highest = math.max(highest, total)
    end
    return highest
  elseif block.t == "Div" or block.t == "BlockQuote" then
    local total = 0
    for _, child in ipairs(block.content) do
      total = total + (height_of(child, width) or 0)
    end
    return total
  end
  return 0
end

list_height = function(list, width, depth)
  local total = 0
  for _, item in ipairs(list.content) do
    for _, child in ipairs(item) do
      if child.t == "BulletList" or child.t == "OrderedList" then
        total = total + list_height(child, width, depth + 1)
      else
        total = total + paragraph_height(text_length(child), width, IIS.text_size, 14 * depth, 0)
      end
    end
  end
  return total
end

-- Size attribute of a picture in cm of the PowerPoint page: "50%" of the
-- reference length, or an absolute length (cm, mm, in, pt). Same rules as
-- size_of() in tools/md2pptx.py.
local UNITS = { cm = 1, mm = 0.1, ["in"] = 2.54, pt = 2.54 / 72 }

local function size_of(value, reference)
  if not value then
    return nil
  end
  local number, unit = tostring(value):match("^%s*([%d%.]+)%s*(%a*%%?)%s*$")
  number = tonumber(number)
  if not number then
    io.stderr:write("layouts: picture size " .. tostring(value) .. " not understood\n")
    return nil
  end
  if unit == "%" then
    return reference * number / 100
  end
  if UNITS[unit] then
    return number * UNITS[unit]
  end
  io.stderr:write("layouts: picture size " .. tostring(value) .. " not understood, use % or cm\n")
  return nil
end

-- height the picture takes when the author set one, nil otherwise
local function explicit_height(block, area_height)
  local image = image_of(block)
  local h = image and size_of(image.attributes.height, area_height)
  if h and caption_of(block) then
    h = h + 1.3
  end
  return h
end

-- width and area_height: cm of the PowerPoint page; height: space left for it
local function fitted_image(block, height, width, area_height)
  local image = image_of(block)
  local caption = caption_of(block)
  local w = size_of(image.attributes.width, width or IIS.text_width)
  local h = size_of(image.attributes.height, area_height or height)
  local box = h or math.max(height - (caption and 1.3 or 0), 1)
  local tex = "{\\centering\\includegraphics[width="
    .. (w and cm(w) or "\\linewidth") .. ",height=" .. cm(box)
    .. ",keepaspectratio]{" .. image.src .. "}\\par"
  if caption then
    tex = tex .. "\\vspace{1.4pt}{\\footnotesize\\color{black!65}"
      .. inlines_to_latex(caption) .. "\\par}"
  end
  return raw(tex .. "}")
end

local function strong_only(block)
  if block.t ~= "Para" and block.t ~= "Plain" then
    return nil
  end
  local inner, seen = pandoc.List(), false
  for _, inline in ipairs(block.content) do
    if inline.t == "Strong" then
      inner:extend(inline.content)
      seen = true
    elseif inline.t == "Space" or inline.t == "SoftBreak" then
      inner:insert(inline)
    else
      return nil
    end
  end
  return seen and inner or nil
end

local function iis_table(block, width)
  local simple = pandoc.utils.to_simple_table(block)
  local rows = pandoc.List()
  if #simple.headers > 0 then
    rows:insert({ head = true, cells = simple.headers })
  end
  for _, row in ipairs(simple.rows) do
    rows:insert({ head = false, cells = row })
  end
  if #rows == 0 then
    return raw("")
  end
  local count = #simple.aligns
  local lengths, sum = {}, 0
  for c = 1, count do
    local longest = 3
    for _, row in ipairs(rows) do
      local cell = row.cells[c]
      if cell then
        longest = math.max(longest, text_length(pandoc.Div(cell)))
      end
    end
    lengths[c] = math.min(math.max(longest, 6), 40)
    sum = sum + lengths[c]
  end
  local natural = (sum * IIS.table_size * 0.6 + count * 16) / 72 * 2.54
  local total = math.min(width, math.max(natural, width / 3))
  local padding = 0.25
  local spec = {}
  for c = 1, count do
    local align = ({ AlignRight = "\\raggedleft", AlignCenter = "\\centering" })[simple.aligns[c]]
      or "\\raggedright"
    spec[c] = string.format(">{%s\\arraybackslash}p{%s}", align,
      cm(total * lengths[c] / sum - 2 * padding))
  end
  local lines = {
    "{\\small\\setlength{\\tabcolsep}{" .. cm(padding) .. "}\\renewcommand{\\arraystretch}{1.5}%",
    "\\arrayrulecolor{white}\\setlength{\\arrayrulewidth}{.5pt}\\rowcolors{2}{IIStableB}{IIStableA}%",
    "\\noindent\\begin{tabular}{" .. table.concat(spec, "|") .. "}",
  }
  for _, row in ipairs(rows) do
    local cells = {}
    for c = 1, count do
      local text = row.cells[c] and blocks_to_latex(row.cells[c]) or ""
      if row.head then
        text = "\\leavevmode\\color{white}\\bfseries " .. text
      end
      cells[c] = text
    end
    lines[#lines + 1] = (row.head and "\\rowcolor{IIStablehead}" or "")
      .. table.concat(cells, " & ") .. " \\\\ \\hline"
  end
  lines[#lines + 1] = "\\end{tabular}\\par"
  local caption = block.caption and block.caption.long or {}
  if #caption > 0 then
    lines[#lines + 1] = "\\vspace{1.4pt}{\\footnotesize\\color{black!65}"
      .. inlines_to_latex(pandoc.utils.blocks_to_inlines(caption)) .. "\\par}"
  end
  lines[#lines + 1] = "}"
  return raw(table.concat(lines, "\n"))
end

local iis_region

local function iis_columns(div, width, height)
  local cols = column_divs(div)
  local count = #cols
  if count == 0 then
    return div
  end
  local inner = (width - IIS.column_gap * (count - 1)) / count
  for _, col in ipairs(cols) do
    col.content = iis_region(col.content, inner, height)
    col.attributes.width = string.format("%.2f%%", inner / width * 100)
  end
  if not div.classes:includes("onlytextwidth") then
    div.classes:insert("onlytextwidth")
  end
  return div
end

-- Quarto wraps the output of code cells in divs.
local function flatten(blocks)
  local out = pandoc.List()
  for _, block in ipairs(blocks) do
    if block.t == "Div" and not is_notes(block) and not is_columns(block)
      and not is_figure_div(block) and not block.classes:includes("column") then
      if not block.classes:includes("hidden") then
        out:extend(flatten(block.content))
      end
    else
      out:insert(block)
    end
  end
  return out
end

-- Pictures share the height that text, code and tables leave free.
iis_region = function(blocks, width, height)
  blocks = flatten(blocks)
  local fixed, flexible, segments, previous_text = 0, 0, 0, false
  for _, block in ipairs(blocks) do
    if not is_notes(block) then
      local h = explicit_height(block, height) or height_of(block, width)
      local is_text = h ~= nil and block.t ~= "CodeBlock" and block.t ~= "Table"
        and not is_columns(block)
      if not (is_text and previous_text) then
        segments = segments + 1
      end
      previous_text = is_text
      if h == nil then
        flexible = flexible + 1
      else
        fixed = fixed + h
      end
    end
  end
  local share = height
  if flexible > 0 then
    share = math.max((height - fixed - IIS.gap * math.max(segments - 1, 0)) / flexible, 2)
  end

  local out = pandoc.List()
  for _, block in ipairs(blocks) do
    local heading = strong_only(block)
    if is_notes(block) then
      out:insert(block)
    elseif image_of(block) then
      out:insert(fitted_image(block, share, width, height))
    elseif is_columns(block) then
      out:insert(iis_columns(block, width, share))
      -- columns end without paragraph skip; PowerPoint keeps a gap to what follows
      out:insert(raw("\\vspace{" .. cm(0.9) .. "}"))
    elseif block.t == "Table" then
      out:insert(iis_table(block, width))
    elseif heading then
      -- paragraph heading: bold, no space below
      out:insert(raw("{\\bfseries " .. inlines_to_latex(heading) .. "\\par}\\vspace{-\\parskip}"))
    else
      out:insert(block)
    end
  end
  return out
end

local function cover_picture(image, left, top, right, bottom)
  local width, height = right - left, bottom - top
  local options = string.format("width=%s,height=%s", cm(width), cm(height))
  local file = io.open(image.src, "rb")
  if file then
    local ok, size = pcall(pandoc.image.size, file:read("a"))
    file:close()
    if ok and size and size.width and size.height and size.height > 0 then
      local ratio = size.width / size.height
      if ratio > width / height then
        options = string.format("height=%s", cm(height))
      else
        options = string.format("width=%s", cm(width))
      end
    end
  end
  return raw(table.concat({
    "\\begin{tikzpicture}[remember picture,overlay]",
    string.format("\\clip ([xshift=%s,yshift=-%s]current page.north west) rectangle "
      .. "([xshift=%s,yshift=-%s]current page.north west);", cm(left), cm(top), cm(right), cm(bottom)),
    string.format("\\node[inner sep=0pt] at ([xshift=%s,yshift=-%s]current page.north west) "
      .. "{\\includegraphics[%s]{%s}};", cm((left + right) / 2), cm((top + bottom) / 2), options,
      image.src),
    "\\end{tikzpicture}%",
  }, "\n"))
end

local function iis_quote(body)
  local paragraphs, notes = pandoc.List(), pandoc.List()
  for _, block in ipairs(body) do
    if is_notes(block) then
      notes:insert(block)
    elseif block.t == "BlockQuote" then
      for _, child in ipairs(block.content) do
        paragraphs:insert(child)
      end
    end
  end
  if #paragraphs == 0 then
    for _, block in ipairs(body) do
      if block.t == "Para" or block.t == "Plain" then
        paragraphs:insert(block)
      end
    end
  end
  if #paragraphs == 0 then
    return nil
  end
  local text = { "{\\bfseries\\fontsize{20.79}{27.44}\\selectfont\\iisragged "
    .. inlines_to_latex(paragraphs[1].content) .. "\\par}" }
  if paragraphs[2] then
    text[#text + 1] = "{\\raggedleft\\bfseries " .. inlines_to_latex(paragraphs[2].content) .. "\\par}"
  end
  if paragraphs[3] then
    text[#text + 1] = "{\\raggedleft " .. inlines_to_latex(paragraphs[3].content) .. "\\par}"
  end
  -- quotation marks of layout 24: two chevrons, 3.5 cm x 2.9 cm at (1.7, 1.3)
  local function chevron(x)
    local t, d, h, y = 1.1, 0.8, 2.9, 1.3
    local points = {
      { x, y }, { x + t, y }, { x + t + d, y + h / 2 }, { x + t, y + h }, { x, y + h },
      { x + d, y + h / 2 },
    }
    local path = {}
    for _, p in ipairs(points) do
      path[#path + 1] = string.format("([xshift=%s,yshift=-%s]current page.north west)",
        cm(p[1]), cm(p[2]))
    end
    return "\\fill[IISdbluegreen] " .. table.concat(path, " -- ") .. " -- cycle;"
  end
  local result = pandoc.List({ raw(table.concat({
    "\\begin{tikzpicture}[remember picture,overlay]",
    chevron(1.7),
    chevron(3.3),
    string.format("\\node[anchor=north west,inner sep=0pt,outer sep=0pt,text width=%s,align=left,"
      .. "text=IISdbluegreen] at ([xshift=%s,yshift=-%s]current page.north west) {\\parskip=0pt",
      cm(IIS.text_width), cm(1.33), cm(5.65)),
    table.concat(text, "\n"),
    "};",
    "\\end{tikzpicture}",
  }, "\n")) })
  result:extend(notes)
  return result
end

local function iis_slide(header, body, layout)
  local options = { "t" }
  local height = IIS.content_bottom - IIS.body_top
  local width = IIS.text_width
  local prefix = pandoc.List()

  if layout == "quote" then
    local quote = iis_quote(body)
    if quote then
      header.content = {}
      header.attributes.subtitle = nil
      return quote, options
    end
  end

  if layout == "blank" then
    header.content = {}
    header.attributes.subtitle = nil
    prefix:insert(raw("\\vspace*{" .. cm(IIS.title_top) .. "}"))
    height = IIS.content_bottom - IIS.title_top
  elseif layout == "content-large" then
    prefix:insert(raw("\\large"
      .. "\\setbeamerfont{itemize/enumerate body}{size=\\large}"
      .. "\\setbeamerfont{itemize/enumerate subbody}{size=\\large}"
      .. "\\setbeamerfont{itemize/enumerate subsubbody}{size=\\large}"))
  elseif layout == "image-left" or layout == "image-right" then
    local image, rest, notes = split_image(body)
    if not image then
      io.stderr:write("layouts: slide with layout " .. layout .. " has no image\n")
    else
      local half = string.format("%.2f%%", IIS.half_width / IIS.text_width * 100)
      local text = pandoc.Div(iis_region(rest, IIS.half_width, height),
        pandoc.Attr("", { "column" }, { width = half }))
      local other
      if layout == "image-left" then
        other = pandoc.Div({ fitted_image(image, height, IIS.half_width, height) },
          pandoc.Attr("", { "column" }, { width = half }))
        body = pandoc.List({ columns({ other, text }, { "columns", "onlytextwidth" }) })
      else
        -- layout 17: the picture fills the right half down to the footer
        other = pandoc.Div({}, pandoc.Attr("", { "column" }, { width = half }))
        body = pandoc.List({ columns({ text, other }, { "columns", "onlytextwidth" }) })
        -- the overlay comes last: in front it would push the text down
        body:insert(cover_picture(image_of(image), IIS.picture_left, 0, IIS.page_width,
          IIS.picture_bottom))
      end
      body:extend(notes)
      prefix:extend(body)
      return prefix, options
    end
  end

  prefix:extend(iis_region(body, width, height))
  return prefix, options
end

-- Date, footer texts and title lines as in the PowerPoint output.
local function iis_meta(meta)
  local lang = meta.lang and pandoc.utils.stringify(meta.lang):sub(1, 2):lower() or "en"
  if lang ~= "de" then
    lang = "en"
  end
  meta["iis-lang"] = lang
  if meta.date then
    local y, m, d = pandoc.utils.stringify(meta.date):match("^(%d%d%d%d)-(%d%d)-(%d%d)$")
    if y then
      local text = lang == "de" and (d .. "." .. m .. "." .. y) or (d .. "/" .. m .. "/" .. y)
      meta.date = pandoc.MetaInlines({ pandoc.Str(text) })
    end
  end
  if meta.footer == nil then
    meta.footer = pandoc.MetaInlines({ pandoc.Str("© Fraunhofer IIS") })
  end
  if meta.classification == nil then
    meta.classification = pandoc.MetaInlines({ pandoc.Str("Restricted") })
  end
  local people = {}
  for _, key in ipairs({ "author", "institute" }) do
    local value = meta[key]
    if value then
      local names = {}
      if pandoc.utils.type(value) == "List" then
        for _, item in ipairs(value) do
          names[#names + 1] = pandoc.utils.stringify(item)
        end
      else
        names[1] = pandoc.utils.stringify(value)
      end
      local text = table.concat(names, ", ")
      if text ~= "" then
        people[#people + 1] = text
      end
    end
  end
  if #people > 0 then
    meta["iis-people"] = pandoc.MetaInlines({ pandoc.Str(table.concat(people, " | ")) })
  end
  return meta
end

-- --------------------------------------------------------------------------

local function slide(header, body, brand)
  local layout = header.attributes.layout or "content"
  header.attributes.layout = nil

  local options
  if brand == "iis" then
    body, options = iis_slide(header, body, layout)
  else
    body, options = fau_slide(header, body, layout)
  end

  local subtitle = header.attributes.subtitle
  header.attributes.subtitle = nil
  if header.attributes.frameoptions then
    options[#options + 1] = header.attributes.frameoptions
  end
  header.attributes.frameoptions = table.concat(options, ",")

  if subtitle and subtitle ~= "" then
    body:insert(1, raw("\\framesubtitle{" .. to_latex(subtitle) .. "}"))
  end

  local result = pandoc.List({ header })
  result:extend(body)
  return result
end

-- --------------------------------------------------------------------------
-- Inline markup shared by all beamer formats
-- --------------------------------------------------------------------------

-- classes for coloured text, also in tools/md2pptx.py
local TEXT_COLOURS = { ok = "179C7D", warn = "F58220", accent = "005B7F" }

-- characters the text fonts do not contain
local SYMBOLS = {
  ["✓"] = "\\ensuremath{\\checkmark}",
  ["✗"] = "\\ensuremath{\\times}",
  ["↑"] = "\\ensuremath{\\uparrow}",
  ["↓"] = "\\ensuremath{\\downarrow}",
  ["→"] = "\\ensuremath{\\rightarrow}",
  ["⟶"] = "\\ensuremath{\\longrightarrow}",
  ["←"] = "\\ensuremath{\\leftarrow}",
}

local function coloured(span)
  for _, class in ipairs(span.classes) do
    local colour = TEXT_COLOURS[class]
    if colour then
      local out = pandoc.List({ pandoc.RawInline("latex", "\\textcolor[HTML]{" .. colour .. "}{") })
      out:extend(span.content)
      out:insert(pandoc.RawInline("latex", "}"))
      return out
    end
  end
  return nil
end

local function symbols(str)
  local out, text, changed = pandoc.List(), "", false
  for _, code in utf8.codes(str.text) do
    local char = utf8.char(code)
    if SYMBOLS[char] then
      if text ~= "" then
        out:insert(pandoc.Str(text))
        text = ""
      end
      out:insert(pandoc.RawInline("latex", SYMBOLS[char]))
      changed = true
    else
      text = text .. char
    end
  end
  if not changed then
    return nil
  end
  if text ~= "" then
    out:insert(pandoc.Str(text))
  end
  return out
end

function Pandoc(doc)
  if not FORMAT:match("beamer") then
    return doc
  end
  doc = doc:walk({ Span = coloured }):walk({ Str = symbols })
  local brand = doc.meta.brand and pandoc.utils.stringify(doc.meta.brand) or ""
  if brand == "iis" then
    doc.meta = iis_meta(doc.meta)
  end
  local blocks, out = doc.blocks, pandoc.List()
  local i = 1
  while i <= #blocks do
    local block = blocks[i]
    if block.t == "Header" and block.level == 2 then
      local body = pandoc.List()
      local j = i + 1
      while j <= #blocks and not (blocks[j].t == "Header" and blocks[j].level <= 2) do
        body:insert(blocks[j])
        j = j + 1
      end
      out:extend(slide(block, body, brand))
      i = j
    else
      out:insert(block)
      i = i + 1
    end
  end
  doc.blocks = out
  return doc
end
