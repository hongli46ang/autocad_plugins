(vl-load-com)

(setq csg:*eps* 1e-9)
(setq csg:*wire-layer* "電氣管線")
(setq csg:*label-layer* "連線群組編號")
(setq csg:*label-prefix* "G")
(setq csg:*min-label-height* 80.0)
(setq csg:*touch-tol-factor* 0.12)
(setq csg:*min-touch-tol* 20.0)
(setq csg:*component-row-tol* 0.0)
(setq csg:*symbol-row-tol* 0.0)
(setq csg:*uf-parent* nil)

(defun csg:trim (s)
  (if s
    (vl-string-trim " \t\r\n" s)
    ""
  )
)

(defun csg:str-eq-ci (a b)
  (= (strcase (csg:trim a)) (strcase (csg:trim b)))
)

(defun csg:pt3 (p / z)
  (setq z (if (> (length p) 2) (nth 2 p) 0.0))
  (list (float (car p)) (float (cadr p)) (float z))
)

(defun csg:dist2 (a b / dx dy dz)
  (setq dx (- (car a) (car b))
        dy (- (cadr a) (cadr b))
        dz (- (caddr a) (caddr b)))
  (+ (* dx dx) (* dy dy) (* dz dz))
)

(defun csg:point-like-p (p)
  (and (listp p)
       (numberp (car p))
       (numberp (cadr p)))
)

(defun csg:segment-like-p (seg)
  (and (listp seg)
       (csg:point-like-p (car seg))
       (csg:point-like-p (cadr seg)))
)

(defun csg:bbox-like-p (bbox)
  (and (listp bbox)
       (csg:point-like-p (car bbox))
       (csg:point-like-p (cadr bbox)))
)

(defun csg:ename-p (val)
  (= (type val) 'ENAME)
)

(defun csg:list-set-nth (lst idx val / i out)
  (setq i 0
        out '())
  (foreach item lst
    (setq out (cons (if (= i idx) val item) out)
          i   (1+ i))
  )
  (reverse out)
)

(defun csg:uf-init (n / i)
  (setq csg:*uf-parent* '()
        i 0)
  (while (< i n)
    (setq csg:*uf-parent* (append csg:*uf-parent* (list i))
          i (1+ i))
  )
)

(defun csg:uf-find (idx / parent root)
  (setq parent (nth idx csg:*uf-parent*))
  (if (= parent idx)
    idx
    (progn
      (setq root (csg:uf-find parent))
      (setq csg:*uf-parent* (csg:list-set-nth csg:*uf-parent* idx root))
      root
    )
  )
)

(defun csg:uf-union (a b / rootA rootB)
  (setq rootA (csg:uf-find a)
        rootB (csg:uf-find b))
  (if (/= rootA rootB)
    (setq csg:*uf-parent* (csg:list-set-nth csg:*uf-parent* rootB rootA))
  )
)

(defun csg:ensure-layer (name color / lay)
  (setq lay (tblsearch "LAYER" name))
  (if lay
    T
    (progn
      (entmake
        (list
          '(0 . "LAYER")
          '(100 . "AcDbSymbolTableRecord")
          '(100 . "AcDbLayerTableRecord")
          (cons 2 name)
          '(70 . 0)
          (cons 62 color)
          '(6 . "Continuous")
        )
      )
      (if (tblsearch "LAYER" name) T nil)
    )
  )
)

(defun csg:layer-color-key (layer / lay color)
  (setq lay   (tblsearch "LAYER" layer)
        color (if lay (cdr (assoc 62 lay)) nil))
  (list 'aci (if color (abs color) 256))
)

(defun csg:entity-color-key (ename / ed trueColor aci layer)
  (setq ed        (if (csg:ename-p ename) (entget ename) nil)
        trueColor (if ed (cdr (assoc 420 ed)) nil)
        aci       (if ed (cdr (assoc 62 ed)) nil)
        layer     (if ed (cdr (assoc 8 ed)) nil))
  (cond
    (trueColor
      (list 'truecolor trueColor)
    )
    ((and aci (/= aci 256))
      (list 'aci (abs aci))
    )
    (layer
      (csg:layer-color-key layer)
    )
    (T
      (list 'aci 256)
    )
  )
)

(defun csg:get-vla-bbox (obj / ret minPt maxPt)
  (if obj
    (progn
      (setq ret (vl-catch-all-apply 'vla-GetBoundingBox (list obj 'minPt 'maxPt)))
      (if (vl-catch-all-error-p ret)
        nil
        (list
          (csg:pt3 (vlax-safearray->list minPt))
          (csg:pt3 (vlax-safearray->list maxPt))
        )
      )
    )
    nil
  )
)

(defun csg:get-bbox (ename / obj)
  (if (csg:ename-p ename)
    (progn
      (setq obj (vlax-ename->vla-object ename))
      (csg:get-vla-bbox obj)
    )
    nil
  )
)

(defun csg:center-from-bbox (bbox / mn mx)
  (setq mn (car bbox)
        mx (cadr bbox))
  (list
    (/ (+ (car mn) (car mx)) 2.0)
    (/ (+ (cadr mn) (cadr mx)) 2.0)
    (/ (+ (caddr mn) (caddr mx)) 2.0)
  )
)

(defun csg:size-from-bbox (bbox / mn mx w h)
  (setq mn (car bbox)
        mx (cadr bbox)
        w  (abs (- (car mx) (car mn)))
        h  (abs (- (cadr mx) (cadr mn))))
  (max w h)
)

(defun csg:quarters-from-bbox (bbox / mn mx cx cy cz)
  (setq mn (car bbox)
        mx (cadr bbox)
        cx (/ (+ (car mn) (car mx)) 2.0)
        cy (/ (+ (cadr mn) (cadr mx)) 2.0)
        cz (/ (+ (caddr mn) (caddr mx)) 2.0))
  (list
    (list cx (cadr mx) cz)
    (list (car mx) cy cz)
    (list cx (cadr mn) cz)
    (list (car mn) cy cz)
  )
)

(defun csg:get-base-point (ename / ed p)
  (if (csg:ename-p ename)
    (progn
      (setq ed (entget ename)
            p  (cdr (assoc 10 ed)))
      (if p
        (csg:pt3 p)
        nil
      )
    )
    nil
  )
)

(defun csg:text-like-type-p (typ)
  (if typ
    (wcmatch (strcase (csg:trim typ)) "TEXT,MTEXT,ATTRIB,ATTDEF,MULTILEADER")
    nil
  )
)

(defun csg:get-effective-block-name (ename / obj val ed)
  (setq obj (vlax-ename->vla-object ename))
  (cond
    ((and obj (vlax-property-available-p obj 'EffectiveName))
      (setq val (vl-catch-all-apply 'vla-get-EffectiveName (list obj)))
      (if (vl-catch-all-error-p val)
        ""
        (csg:trim val)
      )
    )
    ((setq ed (entget ename))
      (csg:trim (cdr (assoc 2 ed)))
    )
    (T "")
  )
)

(defun csg:get-text-from-entity (ename / ed typ obj txt val)
  (setq txt "")
  (if ename
    (progn
      (setq ed  (entget ename)
            typ (cdr (assoc 0 ed)))
      (if (csg:text-like-type-p typ)
        (progn
          (setq obj (vlax-ename->vla-object ename))
          (if obj
            (progn
              (setq val (vl-catch-all-apply 'vla-get-TextString (list obj)))
              (if (not (vl-catch-all-error-p val))
                (setq txt val)
              )
            )
          )
        )
      )
    )
  )
  (csg:trim txt)
)

(defun csg:get-first-nonempty-text-in-ss (ss / i total en txt out)
  (setq i 0
        total (if ss (sslength ss) 0)
        out "")
  (while (and (< i total) (= out ""))
    (setq en  (ssname ss i)
          txt (csg:get-text-from-entity en))
    (if (/= txt "")
      (setq out txt)
    )
    (setq i (1+ i))
  )
  out
)

(defun csg:first-text-entity-in-ss (ss / i total en txt out)
  (setq i 0
        total (if ss (sslength ss) 0)
        out nil)
  (while (and (< i total) (null out))
    (setq en  (ssname ss i)
          txt (csg:get-text-from-entity en))
    (if (/= txt "")
      (setq out en)
    )
    (setq i (1+ i))
  )
  out
)

(defun csg:ss-to-list (ss / i total out)
  (setq i 0
        total (if ss (sslength ss) 0)
        out '())
  (while (< i total)
    (setq out (cons (ssname ss i) out)
          i (1+ i))
  )
  (reverse out)
)

(defun csg:spec-text-token (spec)
  (cond
    ((= (car spec) 'insert)
      (if (> (length spec) 2) (nth 2 spec) "")
    )
    ((= (car spec) 'entity)
      (if (> (length spec) 3) (nth 3 spec) "")
    )
    ((= (car spec) 'text)
      (if (> (length spec) 3) (nth 3 spec) "")
    )
    (T "")
  )
)

(defun csg:entity-has-nearby-text-p (ename token / bbox mn mx size margin p1 p2 ss i total en txt found)
  (if (= (csg:trim token) "")
    T
    (progn
      (setq bbox (csg:get-bbox ename))
      (if bbox
        (progn
          (setq mn (car bbox)
                mx (cadr bbox)
                size (max (abs (- (car mx) (car mn)))
                          (abs (- (cadr mx) (cadr mn))))
                margin (max 50.0 (* size 0.8))
                p1 (list (- (car mn) margin) (- (cadr mn) margin) 0.0)
                p2 (list (+ (car mx) margin) (+ (cadr mx) margin) 0.0)
                ss (ssget "_C" p1 p2 '((0 . "TEXT,MTEXT,ATTRIB,ATTDEF"))))

          (setq i 0
                total (if ss (sslength ss) 0)
                found nil)
          (while (and (< i total) (not found))
            (setq en  (ssname ss i)
                  txt (csg:get-text-from-entity en))
            (if (csg:str-eq-ci token txt)
              (setq found T)
            )
            (setq i (1+ i))
          )
          found
        )
        nil
      )
    )
  )
)

(defun csg:make-symbol-spec (ename legendText / ed typ lyr txt token)
  (setq ed  (entget ename)
        typ (cdr (assoc 0 ed))
        lyr (cdr (assoc 8 ed))
        txt (if (csg:text-like-type-p typ) (csg:get-text-from-entity ename) "")
        token (strcase (csg:trim legendText)))
  (cond
    ((= typ "INSERT")
      (if (= token "")
        (list 'insert (strcase (csg:get-effective-block-name ename)))
        (list 'insert (strcase (csg:get-effective-block-name ename)) token)
      )
    )
    ((and (csg:text-like-type-p typ) (/= txt ""))
      (list 'text
            (strcase (csg:trim typ))
            (strcase (csg:trim lyr))
            (strcase txt))
    )
    (T
      (if (= token "")
        (list 'entity (strcase (csg:trim typ)) (strcase (csg:trim lyr)))
        (list 'entity (strcase (csg:trim typ)) (strcase (csg:trim lyr)) token)
      )
    )
  )
)

(defun csg:symbol-match-p (ename spec / ed typ lyr txt token)
  (setq ed  (entget ename)
        typ (cdr (assoc 0 ed))
        lyr (cdr (assoc 8 ed))
        txt (csg:get-text-from-entity ename)
        token (csg:spec-text-token spec))
  (cond
    ((= (car spec) 'insert)
      (and (= typ "INSERT")
           (csg:str-eq-ci (cadr spec) (csg:get-effective-block-name ename))
           (csg:entity-has-nearby-text-p ename token))
    )
    ((= (car spec) 'text)
      (and (csg:str-eq-ci (cadr spec) typ)
           (csg:str-eq-ci (caddr spec) lyr)
           (csg:str-eq-ci (cadddr spec) txt))
    )
    ((= (car spec) 'entity)
      (and (csg:str-eq-ci (cadr spec) typ)
           (csg:str-eq-ci (caddr spec) lyr)
           (csg:entity-has-nearby-text-p ename token))
    )
    (T nil)
  )
)

(defun csg:symbol-match-relaxed-p (ename spec / ed typ txt token)
  (setq ed  (entget ename)
        typ (cdr (assoc 0 ed))
        txt (csg:get-text-from-entity ename)
        token (csg:spec-text-token spec))
  (cond
    ((= (car spec) 'text)
      (and (csg:str-eq-ci (cadr spec) typ)
           (csg:str-eq-ci (cadddr spec) txt))
    )
    ((= (car spec) 'entity)
      (and (csg:str-eq-ci (cadr spec) typ)
           (csg:entity-has-nearby-text-p ename token))
    )
    (T nil)
  )
)

(defun csg:legend-type-score (typ / t2)
  (setq t2 (strcase (csg:trim typ)))
  (cond
    ((= t2 "INSERT") 0)
    ((= t2 "CIRCLE") 1)
    ((= t2 "ELLIPSE") 2)
    ((csg:text-like-type-p t2) 3)
    ((= t2 "ARC") 4)
    ((wcmatch t2 "LWPOLYLINE,POLYLINE") 5)
    ((= t2 "LINE") 6)
    (T 9)
  )
)

(defun csg:legend-entity-score (ename / typ txt score)
  (setq typ   (cdr (assoc 0 (entget ename)))
        score (csg:legend-type-score typ))
  (if (and (csg:text-like-type-p typ)
           (= (csg:get-text-from-entity ename) ""))
    (+ score 2)
    score
  )
)

(defun csg:pick-legend-from-ss (ss / i total en bestEnt bestScore score)
  (setq i 0
        total (if ss (sslength ss) 0)
        bestEnt nil
        bestScore nil)
  (while (< i total)
    (setq en    (ssname ss i)
          score (csg:legend-entity-score en))
    (if (or (null bestScore) (< score bestScore))
      (progn
        (setq bestEnt en)
        (setq bestScore score)
      )
    )
    (setq i (1+ i))
  )
  bestEnt
)

(defun csg:get-legend-selection (/ ss en txt)
  (prompt "\n1/4 框選/多選要分組的圖示圖例: ")
  (setq ss (ssget))
  (if ss
    (progn
      (setq en  (csg:pick-legend-from-ss ss)
            txt (csg:get-first-nonempty-text-in-ss ss))
      (if en
        (progn
          (if (> (sslength ss) 1)
            (prompt "\n已從框選結果自動判定圖例（優先 Block/圓/文字）。")
          )
          (if (/= (csg:trim txt) "")
            (prompt (strcat "\n圖例文字條件: " txt))
          )
          (list en txt)
        )
        nil
      )
    )
    nil
  )
)

(defun csg:get-area-selection ()
  (prompt "\n2/4 框選/多選要分析的圖示與連線範圍: ")
  (ssget)
)

(defun csg:collect-symbols-by-rule (ss spec rule / i total en out)
  (setq i 0
        total (if ss (sslength ss) 0)
        out '())
  (while (< i total)
    (setq en (ssname ss i))
    (if (and en
             (or (and (= rule 'exact) (csg:symbol-match-p en spec))
                 (and (= rule 'relaxed) (csg:symbol-match-relaxed-p en spec))))
      (setq out (cons en out))
    )
    (setq i (1+ i))
  )
  (reverse out)
)

(defun csg:collect-symbols (ss spec / out token)
  (setq token (csg:spec-text-token spec))
  (setq out (csg:collect-symbols-by-rule ss spec 'exact))
  (if (and (null out) (not (= (car spec) 'insert)))
    (progn
      (setq out (csg:collect-symbols-by-rule ss spec 'relaxed))
      (if out
        (prompt "\n3/4 精準比對為 0，已改用寬鬆比對（忽略圖層）重試。")
      )
    )
  )
  (if (and (null out) (/= (csg:trim token) ""))
    (prompt (strcat "\n3/4 目前已啟用圖例文字條件: " token "，請確認範圍內物件附近有相同文字。"))
  )
  out
)

(defun csg:add-unique-spec (spec specs)
  (if (vl-some '(lambda (x) (equal x spec)) specs)
    specs
    (cons spec specs)
  )
)

(defun csg:make-colored-symbol-spec (baseSpec colorKey)
  (if baseSpec
    (list 'colored baseSpec colorKey)
    nil
  )
)

(defun csg:colored-spec-p (spec)
  (= (car spec) 'colored)
)

(defun csg:colored-spec-base (spec)
  (cadr spec)
)

(defun csg:colored-spec-color (spec)
  (caddr spec)
)

(defun csg:base-symbol-match-spec-p (ename spec)
  (or (csg:symbol-match-p ename spec)
      (and (not (= (car spec) 'insert))
           (csg:symbol-match-relaxed-p ename spec)))
)

(defun csg:make-symbol-spec-for-multi (ename / ed typ txt baseSpec)
  (setq ed  (entget ename)
        typ (cdr (assoc 0 ed))
        txt (csg:get-text-from-entity ename))
  (cond
    ((or (null typ) (csg:wire-type-p typ))
      nil
    )
    ((or (= typ "INSERT")
         (= typ "CIRCLE")
         (= typ "ELLIPSE")
         (csg:text-like-type-p typ))
      (setq baseSpec (csg:make-symbol-spec ename txt))
      (csg:make-colored-symbol-spec baseSpec (csg:entity-color-key ename))
    )
    (T nil)
  )
)

(defun csg:get-sort-legend-specs (/ ss i total en spec specs)
  (prompt "\n1/5 框選/多選適用此排序的圖例（會包含顏色條件）: ")
  (setq ss (ssget))
  (if ss
    (progn
      (setq i 0
            total (sslength ss)
            specs '())
      (while (< i total)
        (setq en   (ssname ss i)
              spec (csg:make-symbol-spec-for-multi en))
        (if spec
          (setq specs (csg:add-unique-spec spec specs))
        )
        (setq i (1+ i))
      )
      (setq specs (reverse specs))
      (if specs
        (progn
          (prompt (strcat "\n已建立可排序圖例條件數量: " (itoa (length specs))))
          specs
        )
        nil
      )
    )
    nil
  )
)

(defun csg:get-panel-point (/ pick en bbox pt)
  (prompt "\n3/5 選取盤體物件，或 Enter 指定盤體位置點: ")
  (setq pick (entsel))
  (cond
    ((and pick (setq en (car pick)) (setq bbox (csg:get-bbox en)))
      (csg:center-from-bbox bbox)
    )
    (T
      (setq pt (getpoint "\n3/5 指定盤體位置點: "))
      (if pt (csg:pt3 pt) nil)
    )
  )
)

(defun csg:symbol-match-any-spec-p (ename specs)
  (vl-some
    '(lambda (spec)
       (cond
         ((csg:colored-spec-p spec)
           (and (equal (csg:entity-color-key ename) (csg:colored-spec-color spec))
                (csg:base-symbol-match-spec-p ename (csg:colored-spec-base spec)))
         )
         (T
           (csg:base-symbol-match-spec-p ename spec)
         )
       )
     )
    specs
  )
)

(defun csg:collect-symbols-by-specs (ss specs / i total en out)
  (setq i 0
        total (if ss (sslength ss) 0)
        out '())
  (while (< i total)
    (setq en (ssname ss i))
    (if (and en (csg:symbol-match-any-spec-p en specs))
      (setq out (cons en out))
    )
    (setq i (1+ i))
  )
  (setq out (reverse out))
  (prompt (strcat "\n4/5 已抓取適用排序的圖示數量: " (itoa (length out))))
  out
)

(defun csg:all-symbol-type-p (typ)
  (if typ
    (wcmatch (strcase (csg:trim typ)) "INSERT,CIRCLE,ELLIPSE")
    nil
  )
)

(defun csg:collect-symbols-by-types (ss typePattern / i total en ed typ out)
  (setq i 0
        total (if ss (sslength ss) 0)
        out '())
  (while (< i total)
    (setq en  (ssname ss i)
          ed  (entget en)
          typ (cdr (assoc 0 ed)))
    (if (and typ (wcmatch (strcase (csg:trim typ)) typePattern))
      (setq out (cons en out))
    )
    (setq i (1+ i))
  )
  (reverse out)
)

(defun csg:collect-all-symbols (ss / inserts simple)
  (setq inserts (csg:collect-symbols-by-types ss "INSERT"))
  (if inserts
    (progn
      (prompt (strcat "\n2/3 已抓取範圍內全部 Block 圖示數量: " (itoa (length inserts))))
      inserts
    )
    (progn
      (setq simple (csg:collect-symbols-by-types ss "CIRCLE,ELLIPSE"))
      (prompt (strcat "\n2/3 未找到 Block，改抓圓/橢圓圖示數量: " (itoa (length simple))))
      simple
    )
  )
)

(defun csg:entity-in-list-p (ename lst)
  (if (vl-some '(lambda (x) (eq x ename)) lst)
    T
    nil
  )
)

(defun csg:wire-type-p (typ)
  (if typ
    (wcmatch (strcase (csg:trim typ)) "LINE,LWPOLYLINE,POLYLINE")
    nil
  )
)

(defun csg:collect-wire-entities (ss symbols / i total en ed typ layer all preferred)
  (setq i 0
        total (if ss (sslength ss) 0)
        all '()
        preferred '())
  (while (< i total)
    (setq en (ssname ss i)
          ed (entget en)
          typ (cdr (assoc 0 ed))
          layer (cdr (assoc 8 ed)))
    (if (and (csg:wire-type-p typ)
             (not (csg:entity-in-list-p en symbols)))
      (progn
        (setq all (cons en all))
        (if (csg:str-eq-ci layer csg:*wire-layer*)
          (setq preferred (cons en preferred))
        )
      )
    )
    (setq i (1+ i))
  )
  (if preferred
    (progn
      (prompt (strcat "\n使用圖層「" csg:*wire-layer* "」上的連線數量: " (itoa (length preferred))))
      (reverse preferred)
    )
    (progn
      (prompt (strcat "\n未找到圖層「" csg:*wire-layer* "」線段，改用範圍內全部 LINE/PLINE 數量: " (itoa (length all))))
      (reverse all)
    )
  )
)

(defun csg:node-from-symbol (idx ename / bbox center size base)
  (setq bbox (csg:get-bbox ename))
  (if bbox
    (progn
      (setq center (csg:center-from-bbox bbox)
            size   (csg:size-from-bbox bbox))
      (list idx ename bbox center size)
    )
    (progn
      (setq base (csg:get-base-point ename))
      (if base
        (list idx ename (list base base) base 0.0)
        nil
      )
    )
  )
)

(defun csg:build-symbol-nodes (symbols / idx out node)
  (setq idx 0
        out '())
  (foreach en symbols
    (setq node (csg:node-from-symbol idx en))
    (if node
      (setq out (cons node out)
            idx (1+ idx))
    )
  )
  (reverse out)
)

(defun csg:avg-symbol-size (nodes / sum cnt sz)
  (setq sum 0.0
        cnt 0)
  (foreach n nodes
    (setq sz (nth 4 n))
    (if (and sz (> sz csg:*eps*))
      (setq sum (+ sum sz)
            cnt (1+ cnt))
    )
  )
  (if (> cnt 0)
    (/ sum cnt)
    0.0
  )
)

(defun csg:lwpolyline-points (ename / ed pts)
  (setq ed (entget ename)
        pts '())
  (foreach d ed
    (if (= (car d) 10)
      (setq pts (append pts (list (csg:pt3 (cdr d)))))
    )
  )
  pts
)

(defun csg:classic-polyline-points (ename / e ed typ pts done)
  (setq e (entnext ename)
        pts '()
        done nil)
  (while (and e (not done))
    (setq ed  (entget e)
          typ (cdr (assoc 0 ed)))
    (cond
      ((= typ "VERTEX")
        (setq pts (append pts (list (csg:pt3 (cdr (assoc 10 ed))))))
      )
      ((= typ "SEQEND")
        (setq done T)
      )
    )
    (setq e (entnext e))
  )
  pts
)

(defun csg:segments-from-points (pts closed / out prev first)
  (setq out '()
        prev nil
        first nil)
  (foreach p pts
    (if (csg:point-like-p p)
      (progn
        (if (null first)
          (setq first p)
        )
        (if prev
          (setq out (append out (list (list prev p))))
        )
        (setq prev p)
      )
    )
  )
  (if (and closed first prev (> (length pts) 2))
    (setq out (append out (list (list prev first))))
  )
  out
)

(defun csg:segments-from-entity (ename / ed typ flag p1 p2 pts closed)
  (setq ed  (entget ename)
        typ (cdr (assoc 0 ed)))
  (cond
    ((= typ "LINE")
      (setq p1 (cdr (assoc 10 ed))
            p2 (cdr (assoc 11 ed)))
      (if (and p1 p2)
        (list (list (csg:pt3 p1) (csg:pt3 p2)))
        '()
      )
    )
    ((= typ "LWPOLYLINE")
      (setq flag   (cdr (assoc 70 ed))
            closed (and flag (= 1 (logand flag 1)))
            pts    (csg:lwpolyline-points ename))
      (csg:segments-from-points pts closed)
    )
    ((= typ "POLYLINE")
      (setq flag   (cdr (assoc 70 ed))
            closed (and flag (= 1 (logand flag 1)))
            pts    (csg:classic-polyline-points ename))
      (csg:segments-from-points pts closed)
    )
    (T '())
  )
)

(defun csg:build-wire-nodes (wireEnts offset / idx out segs)
  (setq idx 0
        out '())
  (foreach en wireEnts
    (setq segs (csg:segments-from-entity en))
    (setq segs (vl-remove-if-not 'csg:segment-like-p segs))
    (if segs
      (setq out (cons (list (+ offset idx) en segs) out)
            idx (1+ idx))
    )
  )
  (reverse out)
)

(defun csg:expanded-bbox (bbox tol / mn mx)
  (setq mn (car bbox)
        mx (cadr bbox))
  (list
    (list (- (car mn) tol) (- (cadr mn) tol) (caddr mn))
    (list (+ (car mx) tol) (+ (cadr mx) tol) (caddr mx))
  )
)

(defun csg:point-in-bbox-p (p bbox / mn mx)
  (setq mn (car bbox)
        mx (cadr bbox))
  (and (>= (car p) (car mn))
       (<= (car p) (car mx))
       (>= (cadr p) (cadr mn))
       (<= (cadr p) (cadr mx)))
)

(defun csg:cross2d (a b c)
  (- (* (- (car b) (car a)) (- (cadr c) (cadr a)))
     (* (- (cadr b) (cadr a)) (- (car c) (car a))))
)

(defun csg:between-p (v a b tol)
  (and (>= v (- (min a b) tol))
       (<= v (+ (max a b) tol)))
)

(defun csg:on-segment-p (p a b tol)
  (and (<= (abs (csg:cross2d a b p)) tol)
       (csg:between-p (car p) (car a) (car b) tol)
       (csg:between-p (cadr p) (cadr a) (cadr b) tol))
)

(defun csg:segments-intersect-p (a b c d / o1 o2 o3 o4 tol)
  (setq tol 1e-7
        o1 (csg:cross2d a b c)
        o2 (csg:cross2d a b d)
        o3 (csg:cross2d c d a)
        o4 (csg:cross2d c d b))
  (or
    (and (< (* o1 o2) 0.0) (< (* o3 o4) 0.0))
    (csg:on-segment-p c a b tol)
    (csg:on-segment-p d a b tol)
    (csg:on-segment-p a c d tol)
    (csg:on-segment-p b c d tol)
  )
)

(defun csg:segment-intersects-bbox-p (seg bbox / a b mn mx p1 p2 p3 p4)
  (if (and (csg:segment-like-p seg) (csg:bbox-like-p bbox))
    (progn
      (setq a  (car seg)
            b  (cadr seg)
            mn (car bbox)
            mx (cadr bbox)
            p1 (list (car mn) (cadr mn) 0.0)
            p2 (list (car mx) (cadr mn) 0.0)
            p3 (list (car mx) (cadr mx) 0.0)
            p4 (list (car mn) (cadr mx) 0.0))
      (or
        (csg:point-in-bbox-p a bbox)
        (csg:point-in-bbox-p b bbox)
        (csg:segments-intersect-p a b p1 p2)
        (csg:segments-intersect-p a b p2 p3)
        (csg:segments-intersect-p a b p3 p4)
        (csg:segments-intersect-p a b p4 p1)
      )
    )
    nil
  )
)

(defun csg:point-segment-dist2 (p a b / vx vy wx wy len2 ratio proj)
  (setq vx (- (car b) (car a))
        vy (- (cadr b) (cadr a))
        wx (- (car p) (car a))
        wy (- (cadr p) (cadr a))
        len2 (+ (* vx vx) (* vy vy)))
  (if (<= len2 csg:*eps*)
    (csg:dist2 p a)
    (progn
      (setq ratio (/ (+ (* wx vx) (* wy vy)) len2))
      (cond
        ((< ratio 0.0) (setq ratio 0.0))
        ((> ratio 1.0) (setq ratio 1.0))
      )
      (setq proj (list (+ (car a) (* ratio vx))
                       (+ (cadr a) (* ratio vy))
                       0.0))
      (csg:dist2 p proj)
    )
  )
)

(defun csg:segment-bboxes-overlap-p (segA segB tol / a b c d)
  (if (and (csg:segment-like-p segA) (csg:segment-like-p segB))
    (progn
      (setq a (car segA)
            b (cadr segA)
            c (car segB)
            d (cadr segB))
      (and
        (<= (- (min (car a) (car b)) tol) (+ (max (car c) (car d)) tol))
        (<= (- (min (car c) (car d)) tol) (+ (max (car a) (car b)) tol))
        (<= (- (min (cadr a) (cadr b)) tol) (+ (max (cadr c) (cadr d)) tol))
        (<= (- (min (cadr c) (cadr d)) tol) (+ (max (cadr a) (cadr b)) tol))
      )
    )
    nil
  )
)

(defun csg:segments-touch-p (segA segB tol / a b c d tol2)
  (if (and (csg:segment-like-p segA) (csg:segment-like-p segB))
    (progn
      (setq a (car segA)
            b (cadr segA)
            c (car segB)
            d (cadr segB)
            tol2 (* tol tol))
      (and
        (csg:segment-bboxes-overlap-p segA segB tol)
        (or
          (csg:segments-intersect-p a b c d)
          (<= (csg:point-segment-dist2 a c d) tol2)
          (<= (csg:point-segment-dist2 b c d) tol2)
          (<= (csg:point-segment-dist2 c a b) tol2)
          (<= (csg:point-segment-dist2 d a b) tol2)
        )
      )
    )
    nil
  )
)

(defun csg:symbol-touches-wire-p (symNode wireNode tol / bbox segs hit)
  (setq bbox (csg:expanded-bbox (nth 2 symNode) tol)
        segs (nth 2 wireNode)
        hit nil)
  (while (and segs (not hit))
    (if (csg:segment-intersects-bbox-p (car segs) bbox)
      (setq hit T)
    )
    (setq segs (cdr segs))
  )
  hit
)

(defun csg:wires-touch-p (wireA wireB tol / segsA segsB hit)
  (setq segsA (nth 2 wireA)
        hit nil)
  (while (and segsA (not hit))
    (setq segsB (nth 2 wireB))
    (while (and segsB (not hit))
      (if (csg:segments-touch-p (car segsA) (car segsB) tol)
        (setq hit T)
      )
      (setq segsB (cdr segsB))
    )
    (setq segsA (cdr segsA))
  )
  hit
)

(defun csg:connect-symbols-and-wires (symNodes wireNodes tol / sym wire a rest)
  (foreach sym symNodes
    (foreach wire wireNodes
      (if (csg:symbol-touches-wire-p sym wire tol)
        (csg:uf-union (car sym) (car wire))
      )
    )
  )

  (setq rest wireNodes)
  (while rest
    (setq a    (car rest)
          rest (cdr rest))
    (foreach b rest
      (if (csg:wires-touch-p a b tol)
        (csg:uf-union (car a) (car b))
      )
    )
  )
)

(defun csg:add-component-stat (root center comps / found x y rec newRec)
  (setq found (assoc root comps)
        x (car center)
        y (cadr center))
  (if found
    (progn
      (setq newRec
            (list
              root
              (+ (cadr found) x)
              (+ (caddr found) y)
              (1+ (cadddr found))
            ))
      (subst newRec found comps)
    )
    (cons (list root x y 1) comps)
  )
)

(defun csg:component-less-p (a b / dy)
  (setq dy (- (caddr a) (caddr b)))
  (if (> (abs dy) csg:*component-row-tol*)
    (> (caddr a) (caddr b))
    (< (cadr a) (cadr b))
  )
)

(defun csg:component-stats (symNodes avgSize / comps root center count)
  (setq comps '())
  (foreach sym symNodes
    (setq root   (csg:uf-find (car sym))
          center (nth 3 sym)
          comps  (csg:add-component-stat root center comps))
  )
  (setq comps
        (mapcar
          '(lambda (rec)
             (setq count (nth 3 rec))
             (list
               (car rec)
               (/ (cadr rec) count)
               (/ (caddr rec) count)
               count
             )
           )
          comps))
  (setq csg:*component-row-tol* (max 1.0 (* avgSize 0.75)))
  (vl-sort comps 'csg:component-less-p)
)

(defun csg:group-map-from-components (components / out idx rec)
  (setq out '()
        idx 1)
  (foreach rec components
    (setq out (cons (cons (car rec) idx) out)
          idx (1+ idx))
  )
  out
)

(defun csg:get-group-num (root groupMap)
  (cdr (assoc root groupMap))
)

(defun csg:label-point (symNode height / bbox mn mx center)
  (setq bbox   (nth 2 symNode)
        mn     (car bbox)
        mx     (cadr bbox)
        center (nth 3 symNode))
  (list (car center) (+ (cadr mx) (* height 0.65)) (caddr center))
)

(defun csg:node-quarter-for-panel (symNode panelPt / bbox q center dx dy)
  (setq bbox   (nth 2 symNode)
        q      (csg:quarters-from-bbox bbox)
        center (nth 3 symNode)
        dx     (- (car panelPt) (car center))
        dy     (- (cadr panelPt) (cadr center)))
  (if (>= (abs dx) (abs dy))
    (if (>= dx 0.0)
      (list (nth 1 q) 'right)
      (list (nth 3 q) 'left)
    )
    (if (>= dy 0.0)
      (list (nth 0 q) 'top)
      (list (nth 2 q) 'bottom)
    )
  )
)

(defun csg:digit-char-p (ch / code)
  (if (and ch (= (strlen ch) 1))
    (progn
      (setq code (ascii ch))
      (and (>= code 48) (<= code 57))
    )
    nil
  )
)

(defun csg:first-number-template (txt / len idx ch start num prefix suffix width)
  (setq txt (csg:trim txt)
        len (strlen txt)
        idx 1
        start nil
        num "")
  (while (and (<= idx len) (null start))
    (setq ch (substr txt idx 1))
    (if (csg:digit-char-p ch)
      (setq start idx)
    )
    (setq idx (1+ idx))
  )
  (if start
    (progn
      (setq idx start)
      (while (and (<= idx len) (csg:digit-char-p (substr txt idx 1)))
        (setq num (strcat num (substr txt idx 1))
              idx (1+ idx))
      )
      (setq prefix (if (> start 1) (substr txt 1 (1- start)) "")
            suffix (if (<= idx len) (substr txt idx) "")
            width  (strlen num))
      (list prefix width suffix)
    )
    (if (= txt "")
      (list csg:*label-prefix* 0 "")
      (list txt 0 "")
    )
  )
)

(defun csg:repeat-char (ch n / out)
  (setq out "")
  (while (> n 0)
    (setq out (strcat out ch)
          n (1- n))
  )
  out
)

(defun csg:zero-pad-int (num width / s)
  (setq s (itoa num))
  (if (> width (strlen s))
    (strcat (csg:repeat-char "0" (- width (strlen s))) s)
    s
  )
)

(defun csg:format-template-label (num template)
  (strcat
    (car template)
    (csg:zero-pad-int num (cadr template))
    (caddr template)
  )
)

(defun csg:text-anchor-point (ename / ed p)
  (setq ed (entget ename)
        p  (cdr (assoc 11 ed)))
  (if (not (csg:point-like-p p))
    (setq p (cdr (assoc 10 ed)))
  )
  (if (csg:point-like-p p)
    (csg:pt3 p)
    nil
  )
)

(defun csg:entity-text-like-p (ename / typ)
  (setq typ (cdr (assoc 0 (entget ename))))
  (csg:text-like-type-p typ)
)

(defun csg:entity-mleader-p (ename / typ)
  (setq typ (cdr (assoc 0 (entget ename))))
  (csg:str-eq-ci typ "MULTILEADER")
)

(defun csg:filter-non-text-entities (ents / out)
  (setq out '())
  (foreach en ents
    (if (or (not (csg:entity-text-like-p en))
            (csg:entity-mleader-p en))
      (setq out (append out (list en)))
    )
  )
  out
)

(defun csg:merged-bbox-from-entities (ents / bbox out)
  (setq out nil)
  (foreach en ents
    (setq bbox (csg:get-bbox en))
    (if bbox
      (setq out
            (if out
              (list
                (list (min (car (car out)) (car (car bbox)))
                      (min (cadr (car out)) (cadr (car bbox)))
                      (min (caddr (car out)) (caddr (car bbox))))
                (list (max (car (cadr out)) (car (cadr bbox)))
                      (max (cadr (cadr out)) (cadr (cadr bbox)))
                      (max (caddr (cadr out)) (caddr (cadr bbox))))
              )
              bbox
            ))
    )
  )
  out
)

(defun csg:extract-arrow-style-data (ename / ed out)
  (setq ed  (entget ename)
        out '())
  (foreach d ed
    (if (member (car d) '(6 48 62 370 420 430 440))
      (setq out (append out (list d)))
    )
  )
  out
)

(defun csg:copy-sample-label-style (/ pick en layer sampleEnts bbox size styleData)
  (prompt "\n2/5 選取一個箭頭範例，或 Enter 使用預設箭頭: ")
  (setq pick (entsel))
  (if pick
    (progn
      (setq en         (car pick)
            sampleEnts (list en)
            bbox       (csg:get-bbox en)
            size       (if bbox (csg:size-from-bbox bbox) nil)
            styleData  (csg:extract-arrow-style-data en)
            layer      csg:*label-layer*)
      (if (and (csg:ename-p en)
               (or (not (csg:entity-text-like-p en))
                   (csg:entity-mleader-p en)))
        (progn
          (prompt "\n套用箭頭範例大小與線型。")
          (list layer size nil nil nil nil nil nil sampleEnts styleData)
        )
        (progn
          (prompt "\n箭頭範例不是可用的非文字物件，改用預設箭頭。")
          nil
        )
      )
    )
    nil
  )
)

(defun csg:default-label-style (height)
  (list csg:*label-layer* height nil 0.0 1 nil nil nil nil nil)
)

(defun csg:label-style-layer (style)
  (if (car style) (car style) csg:*label-layer*)
)

(defun csg:label-style-height (style fallback)
  (if (and (cadr style) (> (cadr style) csg:*eps*))
    (cadr style)
    fallback
  )
)

(defun csg:label-style-template (style)
  (if (nth 5 style)
    (nth 5 style)
    (list csg:*label-prefix* 0 "")
  )
)

(defun csg:label-style-sample-ss (style)
  (nth 6 style)
)

(defun csg:label-style-sample-base (style)
  (nth 7 style)
)

(defun csg:label-style-sample-ents (style)
  (nth 8 style)
)

(defun csg:label-style-dxf-style (style)
  (nth 9 style)
)

(defun csg:make-label (pt label height / data)
  (setq data
        (list
          '(0 . "TEXT")
          (cons 8 csg:*label-layer*)
          (cons 10 pt)
          (cons 11 pt)
          (cons 40 height)
          (cons 1 label)
          '(50 . 0.0)
          '(72 . 1)
          '(73 . 2)
          '(62 . 1)
        ))
  (entmakex data)
)

(defun csg:make-label-with-style (pt label style fallbackHeight / data layer height textStyle rot color)
  (setq layer     (csg:label-style-layer style)
        height    (csg:label-style-height style fallbackHeight)
        textStyle (nth 2 style)
        rot       (if (nth 3 style) (nth 3 style) 0.0)
        color     (nth 4 style)
        data
          (append
            (list
              '(0 . "TEXT")
              (cons 8 layer)
              (cons 10 pt)
              (cons 11 pt)
              (cons 40 height)
              (cons 1 label)
              (cons 50 rot)
              '(72 . 1)
              '(73 . 2)
            )
            (if textStyle (list (cons 7 textStyle)) '())
            (if color (list (cons 62 color)) '())
          ))
  (entmakex data)
)

(defun csg:make-arrow-segment (p1 p2 style / data layer styleData)
  (setq layer     (csg:label-style-layer style)
        styleData (csg:label-style-dxf-style style))
  (setq data
        (append
          (list
            '(0 . "LINE")
            (cons 8 layer)
            (cons 10 p1)
            (cons 11 p2)
          )
          (if styleData styleData (list '(62 . 1)))
        ))
  (entmakex data)
)

(defun csg:side-vector (side)
  (cond
    ((= side 'right)  '(1.0 0.0 0.0))
    ((= side 'left)   '(-1.0 0.0 0.0))
    ((= side 'top)    '(0.0 1.0 0.0))
    ((= side 'bottom) '(0.0 -1.0 0.0))
    (T '(0.0 1.0 0.0))
  )
)

(defun csg:pt-add-scaled (pt vec scale)
  (list
    (+ (car pt) (* (car vec) scale))
    (+ (cadr pt) (* (cadr vec) scale))
    (caddr pt)
  )
)

(defun csg:make-default-arrow (targetPt side style fallbackHeight / height arrowLen headLen headWidth dir perp start wingA wingB made)
  (setq height    (csg:label-style-height style fallbackHeight)
        arrowLen  (max fallbackHeight height)
        headLen   (* arrowLen 0.28)
        headWidth (* arrowLen 0.16)
        dir       (csg:side-vector side)
        perp      (list (- (cadr dir)) (car dir) 0.0)
        start     (csg:pt-add-scaled targetPt dir arrowLen)
        wingA     (csg:pt-add-scaled (csg:pt-add-scaled targetPt dir headLen) perp headWidth)
        wingB     (csg:pt-add-scaled (csg:pt-add-scaled targetPt dir headLen) perp (- headWidth))
        made      0)
  (if (csg:make-arrow-segment start targetPt style)
    (setq made (1+ made))
  )
  (if (csg:make-arrow-segment targetPt wingA style)
    (setq made (1+ made))
  )
  (if (csg:make-arrow-segment targetPt wingB style)
    (setq made (1+ made))
  )
  (> made 0)
)

(defun csg:entities-after (before / e out)
  (setq e (if before (entnext before) (entnext))
        out '())
  (while e
    (setq out (append out (list e))
          e (entnext e))
  )
  out
)

(defun csg:set-entity-layer (ename layer / ed old)
  (setq ed (entget ename))
  (if ed
    (progn
      (setq old (assoc 8 ed))
      (if old
        (entmod (subst (cons 8 layer) old ed))
        (entmod (append ed (list (cons 8 layer))))
      )
    )
  )
)

(defun csg:set-text-string (ename value / obj ret ed old)
  (setq obj (vlax-ename->vla-object ename))
  (if (and obj (vlax-property-available-p obj 'TextString))
    (progn
      (setq ret (vl-catch-all-apply 'vla-put-TextString (list obj value)))
      (not (vl-catch-all-error-p ret))
    )
    (progn
      (setq ed (entget ename)
            old (assoc 1 ed))
      (if old
        (progn
          (entmod (subst (cons 1 value) old ed))
          T
        )
        nil
      )
    )
  )
)

(defun csg:update-first-text-in-entities (ents label / en done)
  (setq done nil)
  (while (and ents (not done))
    (setq en (car ents))
    (if (/= (csg:get-text-from-entity en) "")
      (setq done (csg:set-text-string en label))
    )
    (setq ents (cdr ents))
  )
  done
)

(defun csg:clear-text-in-entities (ents / en)
  (foreach en ents
    (if (/= (csg:get-text-from-entity en) "")
      (csg:set-text-string en "")
    )
  )
)

(defun csg:copy-one-entity-by-points (ename base target / obj copied moved copyEn)
  (setq obj (if (csg:ename-p ename) (vlax-ename->vla-object ename)))
  (if obj
    (progn
      (setq copied (vl-catch-all-apply 'vla-Copy (list obj)))
      (if (vl-catch-all-error-p copied)
        nil
        (progn
          (setq moved
                (vl-catch-all-apply
                  'vla-Move
                  (list copied (vlax-3d-point base) (vlax-3d-point target))
                ))
          (if (vl-catch-all-error-p moved)
            nil
            (progn
              (setq copyEn (vlax-vla-object->ename copied))
              (if (csg:ename-p copyEn) copyEn nil)
            )
          )
        )
      )
    )
    nil
  )
)

(defun csg:copy-sample-entities (ents base target / out copyEn)
  (setq out '())
  (foreach en ents
    (setq copyEn (csg:copy-one-entity-by-points en base target))
    (if copyEn
      (setq out (append out (list copyEn)))
    )
  )
  out
)

(defun csg:copy-sample-arrow-with-style (targetPt side style fallbackHeight / sampleEnts base layer newEnts)
  (setq sampleEnts (csg:label-style-sample-ents style)
        base       (csg:label-style-sample-base style)
        layer      (csg:label-style-layer style))
  (if (and sampleEnts (csg:point-like-p base))
    (progn
      (setq newEnts (csg:copy-sample-entities sampleEnts base targetPt))
      (if newEnts
        (progn
          (foreach en newEnts
            (csg:set-entity-layer en layer)
          )
          (csg:clear-text-in-entities newEnts)
          T
        )
        (csg:make-default-arrow targetPt side style fallbackHeight)
      )
    )
    (csg:make-default-arrow targetPt side style fallbackHeight)
  )
)

(defun csg:delete-existing-labels-on-layer (layer keepEnts / ans ss i en)
  (initget "Yes No")
  (setq ans (getkword (strcat "\n刪除既有「" layer "」圖層物件 [Yes/No] <Yes>: ")))
  (if (or (null ans) (= ans "Yes"))
    (progn
      (setq ss (ssget "_X" (list (cons 8 layer)))
            i 0)
      (if ss
        (while (< i (sslength ss))
          (setq en (ssname ss i))
          (if (not (csg:entity-in-list-p en keepEnts))
            (entdel en)
          )
          (setq i (1+ i))
        )
      )
    )
  )
)

(defun csg:delete-existing-labels ()
  (csg:delete-existing-labels-on-layer csg:*label-layer* nil)
)

(defun csg:add-labels (symNodes groupMap height / sym root num label made)
  (setq made 0)
  (foreach sym symNodes
    (setq root  (csg:uf-find (car sym))
          num   (csg:get-group-num root groupMap)
          label (strcat csg:*label-prefix* (itoa num)))
    (if (csg:make-label (csg:label-point sym height) label height)
      (setq made (1+ made))
    )
  )
  made
)

(defun csg:symbol-less-for-label-p (a b / ca cb dy)
  (setq ca (nth 3 a)
        cb (nth 3 b)
        dy (- (cadr ca) (cadr cb)))
  (if (> (abs dy) csg:*symbol-row-tol*)
    (> (cadr ca) (cadr cb))
    (< (car ca) (car cb))
  )
)

(defun csg:representative-symbol-for-root (root symNodes / matches)
  (setq matches '())
  (foreach sym symNodes
    (if (= (csg:uf-find (car sym)) root)
      (setq matches (cons sym matches))
    )
  )
  (car (vl-sort matches 'csg:symbol-less-for-label-p))
)

(defun csg:representative-symbol-for-root-by-panel (root symNodes panelPt / best bestD d center)
  (foreach sym symNodes
    (if (= (csg:uf-find (car sym)) root)
      (progn
        (setq center (nth 3 sym)
              d      (csg:dist2 center panelPt))
        (if (or (null bestD) (< d bestD))
          (setq best  sym
                bestD d)
        )
      )
    )
  )
  (if best
    best
    (csg:representative-symbol-for-root root symNodes)
  )
)

(defun csg:add-component-arrows (components symNodes style fallbackHeight panelPt / rec root rep quarterInfo target side made)
  (setq made 0)
  (foreach rec components
    (setq root  (car rec)
          rep   (csg:representative-symbol-for-root-by-panel root symNodes panelPt))
    (if rep
      (progn
        (setq quarterInfo (csg:node-quarter-for-panel rep panelPt)
              target      (car quarterInfo)
              side        (cadr quarterInfo))
        (if (csg:make-default-arrow target side style fallbackHeight)
          (setq made (1+ made))
        )
      )
    )
  )
  made
)

(defun csg:report-components (components groupMap / rec)
  (foreach rec components
    (prompt
      (strcat
        "\n"
        csg:*label-prefix*
        (itoa (cdr (assoc (car rec) groupMap)))
        ": "
        (itoa (nth 3 rec))
        " 個圖示"
      )
    )
  )
)

(defun csg:run (/ oldErr oldEcho legendInfo legend legendText spec areaSS symbols symNodes wireEnts wireNodes avgSize tol labelHeight components groupMap labelCount)
  (vl-load-com)
  (setq oldErr  *error*
        oldEcho (getvar "CMDECHO"))

  (defun *error* (msg)
    (setvar "CMDECHO" oldEcho)
    (if oldErr
      (setq *error* oldErr)
      (setq *error* nil)
    )
    (cond
      ((or (= msg "Function cancelled")
           (= msg "quit / exit abort"))
        (prompt "\nAUTOCONNECTGROUPS 已取消。")
      )
      ((and msg (/= msg ""))
        (prompt (strcat "\n錯誤: " msg))
      )
    )
    (princ)
  )

  (setvar "CMDECHO" 0)
  (prompt "\n=== AUTOCONNECTGROUPS 連線圖示分組編號 ===")

  (setq legendInfo (csg:get-legend-selection))
  (if (null legendInfo)
    (prompt "\n未選取圖例，已取消。")
    (progn
      (setq legend     (car legendInfo)
            legendText (cadr legendInfo)
            spec       (csg:make-symbol-spec legend legendText))

      (setq areaSS (csg:get-area-selection))
      (if (null areaSS)
        (prompt "\n未選取分析範圍，已取消。")
        (progn
          (setq symbols (csg:collect-symbols areaSS spec))
          (prompt (strcat "\n3/4 已過濾圖示數量: " (itoa (length symbols))))

          (if (< (length symbols) 1)
            (prompt "\n範圍內沒有符合圖例的圖示，已取消。")
            (progn
              (setq symNodes (csg:build-symbol-nodes symbols))
              (if (< (length symNodes) 1)
                (prompt "\n可用圖示幾何資訊不足，已取消。")
                (progn
                  (setq wireEnts (csg:collect-wire-entities areaSS symbols)
                        wireNodes (csg:build-wire-nodes wireEnts (length symNodes))
                        avgSize (csg:avg-symbol-size symNodes))
                  (if (< (length wireNodes) 1)
                    (prompt "\n範圍內沒有可分析的連線，將每個圖示視為獨立群組。")
                  )

                  (setq tol (max csg:*min-touch-tol* (* avgSize csg:*touch-tol-factor*))
                        labelHeight (max csg:*min-label-height* (* avgSize 0.18)))
                  (prompt (strcat "\n4/4 連線接觸容差: " (rtos tol 2 2)))

                  (csg:uf-init (+ (length symNodes) (length wireNodes)))
                  (csg:connect-symbols-and-wires symNodes wireNodes tol)

                  (setq components (csg:component-stats symNodes avgSize)
                        groupMap   (csg:group-map-from-components components))

                  (if (csg:ensure-layer csg:*label-layer* 1)
                    (progn
                      (csg:delete-existing-labels)
                      (setq labelCount (csg:add-labels symNodes groupMap labelHeight))
                      (prompt
                        (strcat
                          "\n完成分組，共 "
                          (itoa (length components))
                          " 組；已標註 "
                          (itoa labelCount)
                          " 個圖示。"
                        )
                      )
                      (csg:report-components components groupMap)
                    )
                    (prompt (strcat "\n無法建立圖層「" csg:*label-layer* "」，已取消標註。"))
                  )
                )
              )
            )
          )
        )
      )
    )
  )

  (setvar "CMDECHO" oldEcho)
  (if oldErr
    (setq *error* oldErr)
    (setq *error* nil)
  )
  (princ)
)

(defun csg:run-all-symbols-one-label (/ oldErr oldEcho sortSpecs sampleStyle panelPt labelStyle labelLayer sampleKeep areaSS symbols symNodes wireEnts wireNodes avgSize tol fallbackHeight components groupMap arrowCount)
  (vl-load-com)
  (setq oldErr  *error*
        oldEcho (getvar "CMDECHO"))

  (defun *error* (msg)
    (setvar "CMDECHO" oldEcho)
    (if oldErr
      (setq *error* oldErr)
      (setq *error* nil)
    )
    (cond
      ((or (= msg "Function cancelled")
           (= msg "quit / exit abort"))
        (prompt "\nAUTOCONNECTGROUPSALL 已取消。")
      )
      ((and msg (/= msg ""))
        (prompt (strcat "\n錯誤: " msg))
      )
    )
    (princ)
  )

  (setvar "CMDECHO" 0)
  (prompt "\n=== AUTOCONNECTGROUPSALL 全圖示連線群組箭頭標註 ===")

  (setq sortSpecs (csg:get-sort-legend-specs))
  (if (null sortSpecs)
    (prompt "\n未選取可排序圖例，已取消。")
    (progn
      (setq sampleStyle (csg:copy-sample-label-style))

      (setq panelPt (csg:get-panel-point))
      (if (null panelPt)
        (prompt "\n未指定盤體位置，已取消。")
        (progn
          (prompt "\n4/5 框選/多選要分析的全部圖示與連線範圍: ")
          (setq areaSS (ssget))
          (if (null areaSS)
            (prompt "\n未選取分析範圍，已取消。")
            (progn
              (setq symbols (csg:collect-symbols-by-specs areaSS sortSpecs))
              (if (< (length symbols) 1)
                (prompt "\n範圍內沒有符合可排序圖例的圖示，已取消。")
                (progn
                  (setq symNodes (csg:build-symbol-nodes symbols))
                  (if (< (length symNodes) 1)
                    (prompt "\n可用圖示幾何資訊不足，已取消。")
                    (progn
                      (setq wireEnts (csg:collect-wire-entities areaSS symbols)
                            wireNodes (csg:build-wire-nodes wireEnts (length symNodes))
                            avgSize (csg:avg-symbol-size symNodes))
                      (if (< (length wireNodes) 1)
                        (prompt "\n範圍內沒有可分析的連線，將每個圖示視為獨立群組。")
                      )

                      (setq tol (max csg:*min-touch-tol* (* avgSize csg:*touch-tol-factor*))
                            fallbackHeight (max csg:*min-label-height* (* avgSize 0.18))
                            labelStyle (if sampleStyle sampleStyle (csg:default-label-style fallbackHeight))
                            labelLayer (csg:label-style-layer labelStyle)
                            sampleKeep (csg:label-style-sample-ents labelStyle)
                            csg:*symbol-row-tol* (max 1.0 (* avgSize 0.75)))
                      (prompt (strcat "\n5/5 連線接觸容差: " (rtos tol 2 2)))

                      (csg:uf-init (+ (length symNodes) (length wireNodes)))
                      (prompt "\n正在建立連通關係...")
                      (csg:connect-symbols-and-wires symNodes wireNodes tol)
                      (prompt "\n連通關係建立完成。")

                      (setq components (csg:component-stats symNodes avgSize)
                            groupMap   (csg:group-map-from-components components))

                      (if (csg:ensure-layer labelLayer 1)
                        (progn
                          (csg:delete-existing-labels-on-layer labelLayer sampleKeep)
                          (setq arrowCount (csg:add-component-arrows components symNodes labelStyle fallbackHeight panelPt))
                          (prompt
                            (strcat
                              "\n完成分組，共 "
                              (itoa (length components))
                              " 組；已放置 "
                              (itoa arrowCount)
                              " 個箭頭。"
                            )
                          )
                          (csg:report-components components groupMap)
                        )
                        (prompt (strcat "\n無法建立圖層「" labelLayer "」，已取消標註。"))
                      )
                    )
                  )
                )
              )
            )
          )
        )
      )
    )
  )

  (setvar "CMDECHO" oldEcho)
  (if oldErr
    (setq *error* oldErr)
    (setq *error* nil)
  )
  (princ)
)

(setq c:AUTOCONNECTGROUPS nil
      c:AUTOCG nil
      c:AUTOCONNECTGROUPSALL nil
      c:AUTOCGA nil)

(defun c:AUTOCONNECTGROUPS ()
  (csg:run)
)

(defun c:AUTOCG ()
  (csg:run)
)

(defun c:AUTOCONNECTGROUPSALL ()
  (csg:run-all-symbols-one-label)
)

(defun c:AUTOCGA ()
  (csg:run-all-symbols-one-label)
)

(prompt "\nAUTOCONNECTGROUPS / AUTOCG / AUTOCONNECTGROUPSALL / AUTOCGA 載入完成。")
(princ)
