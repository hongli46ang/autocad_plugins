(vl-catch-all-apply 'vl-load-com nil)

(setq mytool:*toolbar-menu-group* "MYTOOLBAR")
(setq mytool:*toolbar-name* "MyToolbar")
(setq mytool:*toolbar-file* "MyToolbar.cuix")

(defun mytool:loaded-file-dir (/ p)
  (setq p
    (cond
      ((and (boundp '*load-truename*) *load-truename*) *load-truename*)
      ((findfile "mytool_startup.lsp") (findfile "mytool_startup.lsp"))
    )
  )
  (if p
    (vl-filename-directory p)
  )
)

(defun mytool:first-existing-file (paths / p out)
  (foreach p paths
    (if (and (not out) p)
      (setq out (findfile p))
    )
  )
  out
)

(defun mytool:toolbar-cui-path (/ dir)
  (setq dir (mytool:loaded-file-dir))
  (mytool:first-existing-file
    (list
      (if dir (strcat dir "\\Resources\\" mytool:*toolbar-file*))
      (if dir (strcat dir "\\" mytool:*toolbar-file*))
      mytool:*toolbar-file*
    )
  )
)

(defun mytool:get-menu-group (groupName / acad groups group)
  (setq acad (vlax-get-acad-object)
        groups (vla-get-MenuGroups acad)
        group (vl-catch-all-apply 'vla-Item (list groups groupName)))
  (if (vl-catch-all-error-p group)
    nil
    group
  )
)

(defun mytool:load-toolbar-cui (cuiPath / acad groups result)
  (if (mytool:get-menu-group mytool:*toolbar-menu-group*)
    T
    (progn
      (setq acad (vlax-get-acad-object)
            groups (vla-get-MenuGroups acad)
            result (vl-catch-all-apply 'vla-Load (list groups cuiPath)))
      (if (vl-catch-all-error-p result)
        (vl-catch-all-apply 'vl-cmdf (list "_.CUILOAD" cuiPath))
      )
      (if (mytool:get-menu-group mytool:*toolbar-menu-group*)
        T
        nil
      )
    )
  )
)

(defun mytool:show-toolbar-by-com (/ group toolbars toolbar result)
  (setq group (mytool:get-menu-group mytool:*toolbar-menu-group*))
  (if group
    (progn
      (setq toolbars (vla-get-Toolbars group)
            toolbar (vl-catch-all-apply 'vla-Item (list toolbars mytool:*toolbar-name*)))
      (if (vl-catch-all-error-p toolbar)
        nil
        (progn
          (setq result (vl-catch-all-apply 'vla-put-Visible (list toolbar :vlax-true)))
          (not (vl-catch-all-error-p result))
        )
      )
    )
    nil
  )
)

(defun mytool:show-toolbar-by-command ()
  (not
    (vl-catch-all-error-p
      (vl-catch-all-apply
        'vl-cmdf
        (list "_.-TOOLBAR" (strcat mytool:*toolbar-menu-group* "." mytool:*toolbar-name*) "_Show")
      )
    )
  )
)

(defun mytool:ensure-toolbar (/ cuiPath loaded shown)
  (setq cuiPath (mytool:toolbar-cui-path))
  (cond
    ((null cuiPath)
      (prompt "\nMyTool: 找不到 MyToolbar.cuix，工具列未載入。"))
    ((not (setq loaded (mytool:load-toolbar-cui cuiPath)))
      (prompt (strcat "\nMyTool: 無法載入工具列自訂檔: " cuiPath)))
    (T
      (setq shown (or (mytool:show-toolbar-by-com)
                      (mytool:show-toolbar-by-command)))
      (if shown
        (prompt "\nMyTool: MyToolbar 工具列已載入並開啟。")
        (prompt "\nMyTool: MyToolbar.cuix 已載入，但工具列未能自動開啟。"))
    )
  )
  (princ)
)

(mytool:ensure-toolbar)
(princ)
