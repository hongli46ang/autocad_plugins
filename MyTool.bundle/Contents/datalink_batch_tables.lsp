(vl-load-com)

;; Backward-compatible loader. The maintained file is singular:
;; datalink_batch_table.lsp
(setq dlbt:*compat-loader* (findfile "datalink_batch_table.lsp"))
(if dlbt:*compat-loader*
  (load dlbt:*compat-loader*)
  (princ "\n找不到 datalink_batch_table.lsp；請改載入 MyTool.bundle/Contents/datalink_batch_table.lsp。")
)
(princ)
