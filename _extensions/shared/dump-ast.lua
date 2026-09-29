-- Writes the final document AST as JSON next to the output file.
-- tools/md2pptx.py builds the branded PPTX from this file.
--
-- Quarto removes generated figure files once the PPTX is written, so image
-- contents are embedded as data URIs in the dumped copy. The document handed
-- back to pandoc is left untouched.

local function embed(img)
  if img.src:match("^data:") then
    return nil
  end
  local ok, mime, contents = pcall(pandoc.mediabag.fetch, img.src)
  if not ok or not contents then
    io.stderr:write("dump-ast: cannot read image " .. img.src .. "\n")
    return nil
  end
  img.attributes["source"] = img.src
  img.src = "data:" .. (mime or "application/octet-stream") .. ";base64,"
    .. quarto.base64.encode(contents)
  return img
end

function Pandoc(doc)
  local out = PANDOC_STATE.output_file
  if not out then
    return doc
  end
  local copy = doc:walk({ Image = embed })
  local f = assert(io.open(out .. ".ast.json", "w"))
  f:write(pandoc.write(copy, "json"))
  f:close()
  return doc
end
