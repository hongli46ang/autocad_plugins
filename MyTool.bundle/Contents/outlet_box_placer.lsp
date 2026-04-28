(vl-load-com)

(setq obp:*spacing* 5000.0)
(setq obp:*eps* 1e-9)
(setq obp:*connect-tol* 1.0)
(setq obp:*center-pair-tol* 1000.0)
(setq obp:*curve-type-pattern* "LINE,LWPOLYLINE,POLYLINE,ARC,SPLINE,ELLIPSE")

(defun obp:ename-p (val)
  (= (type val) 'ENAME)
)

(defun obp:pt3 (p / z)
  (if (and (listp p) (numberp (car p)) (numberp (cadr p)))
    (progn
      (setq z (if (and (cddr p) (numberp (caddr p))) (caddr p) 0.0))
      (list (float (car p)) (float (cadr p)) (float z))
    )
    nil
  )
)

(defun obp:ename->vla (ename / obj)
  (if (obp:ename-p ename)
    (progn
      (setq obj (vl-catch-all-apply 'vlax-ename->vla-object (list ename)))
      (if (vl-catch-all-error-p obj) nil obj)
    )
    nil
  )
)

(defun obp:ss-to-list (ss / i total out)
  (setq i 0
        total (if ss (sslength ss) 0)
        out '())
  (while (< i total)
    (setq out (append out (list (ssname ss i)))
          i   (1+ i))
  )
  out
)

(defun obp:get-vla-bbox (obj / ret minPt maxPt)
  (if obj
    (progn
      (setq ret (vl-catch-all-apply 'vla-GetBoundingBox (list obj 'minPt 'maxPt)))
      (if (vl-catch-all-error-p ret)
        nil
        (list
          (obp:pt3 (vlax-safearray->list minPt))
          (obp:pt3 (vlax-safearray->list maxPt))
        )
      )
    )
    nil
  )
)

(defun obp:get-bbox (ename)
  (obp:get-vla-bbox (obp:ename->vla ename))
)

(defun obp:merge-bbox (a b / amin amax bmin bmax)
  (cond
    ((null a) b)
    ((null b) a)
    (T
      (setq amin (car a)
            amax (cadr a)
            bmin (car b)
            bmax (cadr b))
      (list
        (list
          (min (car amin) (car bmin))
          (min (cadr amin) (cadr bmin))
          (min (caddr amin) (caddr bmin))
        )
        (list
          (max (car amax) (car bmax))
          (max (cadr amax) (cadr bmax))
          (max (caddr amax) (caddr bmax))
        )
      )
    )
  )
)

(defun obp:get-entities-bbox (ents / bbox en)
  (setq bbox nil)
  (foreach en ents
    (setq bbox (obp:merge-bbox bbox (obp:get-bbox en)))
  )
  bbox
)

(defun obp:center-from-bbox (bbox / mn mx)
  (if bbox
    (progn
      (setq mn (car bbox)
            mx (cadr bbox))
      (list
        (/ (+ (car mn) (car mx)) 2.0)
        (/ (+ (cadr mn) (cadr mx)) 2.0)
        (/ (+ (caddr mn) (caddr mx)) 2.0)
      )
    )
    nil
  )
)

(defun obp:get-sample-selection (/ ss ents bbox base)
  (prompt "\n1/6 框選/多選出線盒圖例: ")
  (setq ss (ssget))
  (if ss
    (progn
      (setq ents (obp:ss-to-list ss)
            bbox (obp:get-entities-bbox ents)
            base (obp:center-from-bbox bbox))
      (if (and ents base)
        (progn
          (prompt "\n已取得出線盒圖例；將以圖例外框中心作為放置中心點。")
          (list ents base)
        )
        nil
      )
    )
    nil
  )
)

(defun obp:curve-type-p (typ)
  (and typ (wcmatch (strcase typ) obp:*curve-type-pattern*))
)

(defun obp:curve-entity-p (ename / ed)
  (setq ed (if (obp:ename-p ename) (entget ename) nil))
  (and ed (obp:curve-type-p (cdr (assoc 0 ed))))
)

(defun obp:curve-class (typ / t2)
  (setq t2 (if typ (strcase typ) ""))
  (cond
    ((wcmatch t2 "LWPOLYLINE,POLYLINE") "PLINE")
    (T t2)
  )
)

(defun obp:layer-color-key (layer / lay color)
  (setq lay   (tblsearch "LAYER" layer)
        color (if lay (cdr (assoc 62 lay)) nil))
  (list 'aci (if color (abs color) 256))
)

(defun obp:entity-color-key (ename / ed trueColor aci layer)
  (setq ed        (if (obp:ename-p ename) (entget ename) nil)
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
      (obp:layer-color-key layer)
    )
    (T
      (list 'aci 256)
    )
  )
)

(defun obp:entity-linetype-key (ename / ed ltype)
  (setq ed    (if (obp:ename-p ename) (entget ename) nil)
        ltype (if ed (cdr (assoc 6 ed)) nil))
  (strcase (if ltype ltype "BYLAYER"))
)

(defun obp:target-spec-from-entity (ename / ed typ layer)
  (setq ed (if (obp:ename-p ename) (entget ename) nil))
  (if (and ed (obp:curve-type-p (cdr (assoc 0 ed))))
    (progn
      (setq typ   (cdr (assoc 0 ed))
            layer (cdr (assoc 8 ed)))
      (list
        (obp:curve-class typ)
        (strcase (if layer layer "0"))
        (obp:entity-color-key ename)
        (obp:entity-linetype-key ename)
      )
    )
    nil
  )
)

(defun obp:add-unique-spec (spec specs / found item)
  (setq found nil)
  (foreach item specs
    (if (equal item spec)
      (setq found T)
    )
  )
  (if found specs (append specs (list spec)))
)

(defun obp:get-target-legend-specs (/ ss i total en spec specs)
  (prompt "\n2/6 框選/多選適用的線/曲線圖例 (LINE/PLINE/ARC/SPLINE/ELLIPSE): ")
  (setq ss (ssget '((0 . "LINE,LWPOLYLINE,POLYLINE,ARC,SPLINE,ELLIPSE"))))
  (setq i 0
        total (if ss (sslength ss) 0)
        specs '())
  (while (< i total)
    (setq en   (ssname ss i)
          spec (obp:target-spec-from-entity en))
    (if spec
      (setq specs (obp:add-unique-spec spec specs))
    )
    (setq i (1+ i))
  )
  (if specs
    (progn
      (prompt (strcat "\n已取得適用物件圖例規則: " (itoa (length specs)) " 種。"))
      specs
    )
    nil
  )
)

(defun obp:get-area-selection ()
  (prompt "\n5/6 框選/多選要放置出線盒的範圍: ")
  (ssget)
)

(defun obp:matches-any-target-spec-p (ename specs / spec target found)
  (setq target (obp:target-spec-from-entity ename)
        found nil)
  (foreach spec specs
    (if (and (not found) target (equal target spec))
      (setq found T)
    )
  )
  found
)

(defun obp:filter-area-by-specs (ss specs / i total en out rejected nonCurve)
  (setq i 0
        total (if ss (sslength ss) 0)
        out '()
        rejected 0
        nonCurve 0)
  (while (< i total)
    (setq en (ssname ss i))
    (cond
      ((not (obp:curve-entity-p en))
        (setq nonCurve (1+ nonCurve))
      )
      ((obp:matches-any-target-spec-p en specs)
        (setq out (cons en out))
      )
      (T
        (setq rejected (1+ rejected))
      )
    )
    (setq i (1+ i))
  )
  (list (reverse out) rejected nonCurve)
)

(defun obp:curve-start-point (ename / p)
  (setq p (vl-catch-all-apply 'vlax-curve-getStartPoint (list ename)))
  (if (vl-catch-all-error-p p) nil (obp:pt3 p))
)

(defun obp:curve-end-point (ename / p)
  (setq p (vl-catch-all-apply 'vlax-curve-getEndPoint (list ename)))
  (if (vl-catch-all-error-p p) nil (obp:pt3 p))
)

(defun obp:curve-length (ename / endParam len)
  (setq endParam (vl-catch-all-apply 'vlax-curve-getEndParam (list ename)))
  (if (vl-catch-all-error-p endParam)
    nil
    (progn
      (setq len (vl-catch-all-apply 'vlax-curve-getDistAtParam (list ename endParam)))
      (if (vl-catch-all-error-p len) nil len)
    )
  )
)

(defun obp:curve-point-at-dist (ename dist / p)
  (setq p (vl-catch-all-apply 'vlax-curve-getPointAtDist (list ename dist)))
  (if (vl-catch-all-error-p p) nil (obp:pt3 p))
)

(defun obp:curve-closed-p (ename / obj closed sp ep)
  (setq obj (obp:ename->vla ename))
  (cond
    ((and obj (vlax-property-available-p obj 'Closed))
      (setq closed (vl-catch-all-apply 'vla-get-Closed (list obj)))
      (if (vl-catch-all-error-p closed)
        nil
        (= closed :vlax-true)
      )
    )
    (T
      (setq sp (obp:curve-start-point ename)
            ep (obp:curve-end-point ename))
      (and sp ep (<= (distance sp ep) obp:*eps*))
    )
  )
)

(defun obp:copy-one-entity-by-points (ename base target / obj copied moved copyEn)
  (setq obj (obp:ename->vla ename))
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
              (if (obp:ename-p copyEn) copyEn nil)
            )
          )
        )
      )
    )
    nil
  )
)

(defun obp:copy-sample-entities (ents base target / out copyEn)
  (setq out '())
  (foreach en ents
    (setq copyEn (obp:copy-one-entity-by-points en base target))
    (if copyEn
      (setq out (append out (list copyEn)))
    )
  )
  out
)

(defun obp:place-on-curve (ename sampleEnts base spacing startRef / len dist curveDist pt newEnts created reverse)
  (cond
    ((obp:curve-closed-p ename)
      (list 0 'closed)
    )
    ((or (null (setq len (obp:curve-length ename)))
         (<= len obp:*eps*))
      (list 0 'invalid)
    )
    (T
      (setq reverse (obp:reverse-for-start-p (obp:curve-start-point ename) (obp:curve-end-point ename) startRef)
            dist 0.0
            created 0)
      (while (<= dist len)
        (setq curveDist (if reverse (- len dist) dist)
              pt        (obp:curve-point-at-dist ename curveDist))
        (if pt
          (progn
            (setq newEnts (obp:copy-sample-entities sampleEnts base pt))
            (if newEnts
              (setq created (1+ created))
            )
          )
        )
        (setq dist (+ dist spacing))
      )
      (if (> created 0)
        (list created 'ok)
        (list 0 'copy-failed)
      )
    )
  )
)

(defun obp:get-place-mode (/ ans)
  (initget "Normal Connect Center")
  (setq ans (getkword "\n3/6 放置模式 [Normal=各自計算/Connect=相連累計/Center=雙線中線] <Normal>: "))
  (cond
    ((= ans "Connect") 'connect)
    ((= ans "Center") 'center)
    (T 'normal)
  )
)

(defun obp:get-start-reference (/ pt)
  (setq pt (getpoint "\n4/6 指定盤位置作為放置起算端，直接 Enter 則由左邊/上面開始: "))
  (if pt
    (progn
      (prompt "\n已指定盤位置；每條路徑會由離盤最近的一端開始。")
      (list 'panel (obp:pt3 pt))
    )
    (progn
      (prompt "\n未指定盤位置；每條路徑會由左邊優先、同 X 時上面優先的一端開始。")
      (list 'default nil)
    )
  )
)

(defun obp:left-top-before-p (a b)
  (cond
    ((< (car a) (- (car b) obp:*eps*))
      T
    )
    ((> (car a) (+ (car b) obp:*eps*))
      nil
    )
    (T
      (> (cadr a) (cadr b))
    )
  )
)

(defun obp:reverse-for-start-p (startPt endPt startRef / mode panel dStart dEnd)
  (setq mode (if startRef (car startRef) 'default))
  (cond
    ((or (null startPt) (null endPt))
      nil
    )
    ((= mode 'panel)
      (setq panel  (cadr startRef)
            dStart (distance panel startPt)
            dEnd   (distance panel endPt))
      (< dEnd dStart)
    )
    (T
      (obp:left-top-before-p endPt startPt)
    )
  )
)

(defun obp:get-center-pair-tol (/ val)
  (setq val
        (getreal
          (strcat
            "\n雙線中線模式端點配對最大距離 <"
            (rtos obp:*center-pair-tol* 2 1)
            ">: "
          )
        ))
  (if (and val (> val 0.0))
    val
    obp:*center-pair-tol*
  )
)

(defun obp:pt-close-p (a b)
  (and a b (<= (distance a b) obp:*connect-tol*))
)

(defun obp:remove-chain-item (target items / out item removed)
  (setq out '()
        removed nil)
  (foreach item items
    (if (and (not removed) (equal (car item) (car target)))
      (setq removed T)
      (setq out (append out (list item)))
    )
  )
  out
)

(defun obp:curve-chain-item (ename allowClosed / sp ep len)
  (cond
    ((and (not allowClosed) (obp:curve-closed-p ename))
      (list nil 'closed)
    )
    ((or (null (setq len (obp:curve-length ename)))
         (<= len obp:*eps*))
      (list nil 'invalid)
    )
    (T
      (setq sp (obp:curve-start-point ename)
            ep (obp:curve-end-point ename))
      (if (and sp ep)
        (list (list ename sp ep len) 'ok)
        (list nil 'invalid)
      )
    )
  )
)

(defun obp:collect-chain-items (ents allowClosed / items skippedClosed skippedInvalid made status en)
  (setq items '()
        skippedClosed 0
        skippedInvalid 0)
  (foreach en ents
    (setq made   (obp:curve-chain-item en allowClosed)
          status (cadr made))
    (cond
      ((= status 'ok)
        (setq items (append items (list (car made))))
      )
      ((= status 'closed)
        (setq skippedClosed (1+ skippedClosed))
      )
      (T
        (setq skippedInvalid (1+ skippedInvalid))
      )
    )
  )
  (list items skippedClosed skippedInvalid)
)

(defun obp:chain-piece (item forward)
  (list (car item) forward (cadddr item))
)

(defun obp:build-one-chain (items / first chain head tail remaining changed item hit piece action newHead newTail)
  (setq first     (car items)
        chain     (list (obp:chain-piece first T))
        head      (cadr first)
        tail      (caddr first)
        remaining (cdr items)
        changed   T)
  (while changed
    (setq changed nil
          hit nil
          piece nil
          action nil)
    (foreach item remaining
      (if (null hit)
        (cond
          ((obp:pt-close-p (cadr item) tail)
            (setq hit item
                  piece (obp:chain-piece item T)
                  action 'append
                  newTail (caddr item))
          )
          ((obp:pt-close-p (caddr item) tail)
            (setq hit item
                  piece (obp:chain-piece item nil)
                  action 'append
                  newTail (cadr item))
          )
          ((obp:pt-close-p (caddr item) head)
            (setq hit item
                  piece (obp:chain-piece item T)
                  action 'prepend
                  newHead (cadr item))
          )
          ((obp:pt-close-p (cadr item) head)
            (setq hit item
                  piece (obp:chain-piece item nil)
                  action 'prepend
                  newHead (caddr item))
          )
        )
      )
    )
    (if hit
      (progn
        (if (= action 'append)
          (setq chain (append chain (list piece))
                tail  newTail)
          (setq chain (cons piece chain)
                head  newHead)
        )
        (setq remaining (obp:remove-chain-item hit remaining)
              changed T)
      )
    )
  )
  (list chain remaining)
)

(defun obp:build-chains (items / remaining chains result)
  (setq remaining items
        chains '())
  (while remaining
    (setq result    (obp:build-one-chain remaining)
          chains    (append chains (list (car result)))
          remaining (cadr result))
  )
  chains
)

(defun obp:chain-length (chain / total piece)
  (setq total 0.0)
  (foreach piece chain
    (setq total (+ total (caddr piece)))
  )
  total
)

(defun obp:chain-point-at-dist (chain dist / rem piece len curveDist pt)
  (setq rem dist
        pt nil)
  (while (and chain (null pt))
    (setq piece (car chain)
          len   (caddr piece))
    (if (<= rem (+ len obp:*eps*))
      (progn
        (cond
          ((< rem 0.0)
            (setq rem 0.0)
          )
          ((> rem len)
            (setq rem len)
          )
        )
        (setq curveDist (if (cadr piece) rem (- len rem))
              pt        (obp:curve-point-at-dist (car piece) curveDist))
      )
      (setq rem   (- rem len)
            chain (cdr chain))
    )
  )
  pt
)

(defun obp:chain-start-point (chain)
  (obp:chain-point-at-dist chain 0.0)
)

(defun obp:chain-end-point (chain)
  (obp:chain-point-at-dist chain (obp:chain-length chain))
)

(defun obp:chain-id (chain)
  (if chain (caar chain) nil)
)

(defun obp:remove-chain (target chains / out chain removed targetId)
  (setq out '()
        removed nil
        targetId (obp:chain-id target))
  (foreach chain chains
    (if (and (not removed) (equal (obp:chain-id chain) targetId))
      (setq removed T)
      (setq out (append out (list chain)))
    )
  )
  out
)

(defun obp:center-pt (a b)
  (list
    (/ (+ (car a) (car b)) 2.0)
    (/ (+ (cadr a) (cadr b)) 2.0)
    (/ (+ (caddr a) (caddr b)) 2.0)
  )
)

(defun obp:chain-pair-info (a b / as ae bs be dss dee dse des sameMax sameScore revMax revScore)
  (setq as (obp:chain-start-point a)
        ae (obp:chain-end-point a)
        bs (obp:chain-start-point b)
        be (obp:chain-end-point b))
  (if (and as ae bs be)
    (progn
      (setq dss (distance as bs)
            dee (distance ae be)
            dse (distance as be)
            des (distance ae bs)
            sameMax (max dss dee)
            sameScore (+ dss dee)
            revMax (max dse des)
            revScore (+ dse des))
      (if (or (< revMax sameMax)
              (and (= revMax sameMax) (< revScore sameScore)))
        (list b T revMax revScore)
        (list b nil sameMax sameScore)
      )
    )
    nil
  )
)

(defun obp:best-chain-pair (chain candidates / best info)
  (setq best nil)
  (foreach candidate candidates
    (setq info (obp:chain-pair-info chain candidate))
    (if (and info
             (or (null best)
                 (< (caddr info) (caddr best))
                 (and (= (caddr info) (caddr best))
                      (< (cadddr info) (cadddr best)))))
      (setq best info)
    )
  )
  best
)

(defun obp:pair-centerline-chains (chains tol / remaining pairs unpaired chain best)
  (setq remaining chains
        pairs '()
        unpaired '())
  (while remaining
    (setq chain     (car remaining)
          remaining (cdr remaining)
          best      (obp:best-chain-pair chain remaining))
    (if (and best (<= (caddr best) tol))
      (progn
        (setq pairs (append pairs (list (list chain (car best) (cadr best))))
              remaining (obp:remove-chain (car best) remaining))
      )
      (setq unpaired (append unpaired (list chain)))
    )
  )
  (list pairs unpaired)
)

(defun obp:centerline-pair-length (pair / lenA lenB)
  (setq lenA (obp:chain-length (car pair))
        lenB (obp:chain-length (cadr pair)))
  (/ (+ lenA lenB) 2.0)
)

(defun obp:centerline-pair-point-at-dist (pair dist / a b reversed lenA lenB len ratio pa pb)
  (setq a        (car pair)
        b        (cadr pair)
        reversed (caddr pair)
        lenA     (obp:chain-length a)
        lenB     (obp:chain-length b)
        len      (/ (+ lenA lenB) 2.0))
  (if (<= len obp:*eps*)
    nil
    (progn
      (setq ratio (/ dist len))
      (cond
        ((< ratio 0.0)
          (setq ratio 0.0)
        )
        ((> ratio 1.0)
          (setq ratio 1.0)
        )
      )
      (setq pa (obp:chain-point-at-dist a (* ratio lenA))
            pb (obp:chain-point-at-dist b (if reversed (* (- 1.0 ratio) lenB) (* ratio lenB))))
      (if (and pa pb)
        (obp:center-pt pa pb)
        nil
      )
    )
  )
)

(defun obp:place-on-centerline-pair (pair sampleEnts base spacing startRef / len dist centerDist pt newEnts created reverse)
  (setq len (obp:centerline-pair-length pair))
  (if (<= len obp:*eps*)
    (list 0 'invalid)
    (progn
      (setq reverse (obp:reverse-for-start-p
                      (obp:centerline-pair-point-at-dist pair 0.0)
                      (obp:centerline-pair-point-at-dist pair len)
                      startRef)
            dist 0.0
            created 0)
      (while (<= dist len)
        (setq centerDist (if reverse (- len dist) dist)
              pt         (obp:centerline-pair-point-at-dist pair centerDist))
        (if pt
          (progn
            (setq newEnts (obp:copy-sample-entities sampleEnts base pt))
            (if newEnts
              (setq created (1+ created))
            )
          )
        )
        (setq dist (+ dist spacing))
      )
      (if (> created 0)
        (list created 'ok)
        (list 0 'copy-failed)
      )
    )
  )
)

(defun obp:self-centerline-chain-p (chain tol / sp ep)
  (setq sp (obp:chain-start-point chain)
        ep (obp:chain-end-point chain))
  (and sp ep (<= (distance sp ep) tol))
)

(defun obp:self-centerline-length (chain)
  (/ (obp:chain-length chain) 2.0)
)

(defun obp:self-centerline-point-at-dist (chain dist / total len pa pb)
  (setq total (obp:chain-length chain)
        len   (/ total 2.0))
  (if (<= len obp:*eps*)
    nil
    (progn
      (cond
        ((< dist 0.0)
          (setq dist 0.0)
        )
        ((> dist len)
          (setq dist len)
        )
      )
      (setq pa (obp:chain-point-at-dist chain dist)
            pb (obp:chain-point-at-dist chain (- total dist)))
      (if (and pa pb)
        (obp:center-pt pa pb)
        nil
      )
    )
  )
)

(defun obp:place-on-self-centerline (chain sampleEnts base spacing startRef / len dist centerDist pt newEnts created reverse)
  (setq len (obp:self-centerline-length chain))
  (if (<= len obp:*eps*)
    (list 0 'invalid)
    (progn
      (setq reverse (obp:reverse-for-start-p
                      (obp:self-centerline-point-at-dist chain 0.0)
                      (obp:self-centerline-point-at-dist chain len)
                      startRef)
            dist 0.0
            created 0)
      (while (<= dist len)
        (setq centerDist (if reverse (- len dist) dist)
              pt         (obp:self-centerline-point-at-dist chain centerDist))
        (if pt
          (progn
            (setq newEnts (obp:copy-sample-entities sampleEnts base pt))
            (if newEnts
              (setq created (1+ created))
            )
          )
        )
        (setq dist (+ dist spacing))
      )
      (if (> created 0)
        (list created 'ok)
        (list 0 'copy-failed)
      )
    )
  )
)

(defun obp:place-on-chain (chain sampleEnts base spacing startRef / len dist chainDist pt newEnts created reverse)
  (setq len (obp:chain-length chain))
  (if (<= len obp:*eps*)
    (list 0 'invalid)
    (progn
      (setq reverse (obp:reverse-for-start-p (obp:chain-start-point chain) (obp:chain-end-point chain) startRef)
            dist 0.0
            created 0)
      (while (<= dist len)
        (setq chainDist (if reverse (- len dist) dist)
              pt        (obp:chain-point-at-dist chain chainDist))
        (if pt
          (progn
            (setq newEnts (obp:copy-sample-entities sampleEnts base pt))
            (if newEnts
              (setq created (1+ created))
            )
          )
        )
        (setq dist (+ dist spacing))
      )
      (if (> created 0)
        (list created 'ok)
        (list 0 'copy-failed)
      )
    )
  )
)

(defun obp:place-independent-curves (pathEnts sampleEnts base startRef / placed processed skippedClosed skippedInvalid skippedCopy en result)
  (setq placed 0
        processed 0
        skippedClosed 0
        skippedInvalid 0
        skippedCopy 0)
  (foreach en pathEnts
    (setq result (obp:place-on-curve en sampleEnts base obp:*spacing* startRef))
    (cond
      ((= (cadr result) 'ok)
        (setq placed    (+ placed (car result))
              processed (1+ processed))
      )
      ((= (cadr result) 'closed)
        (setq skippedClosed (1+ skippedClosed))
      )
      ((= (cadr result) 'copy-failed)
        (setq skippedCopy (1+ skippedCopy))
      )
      (T
        (setq skippedInvalid (1+ skippedInvalid))
      )
    )
  )
  (list placed processed skippedClosed skippedInvalid skippedCopy 0)
)

(defun obp:place-chained-curves (pathEnts sampleEnts base startRef / collected items chains chain result
                                          placed processed skippedClosed skippedInvalid skippedCopy)
  (setq collected      (obp:collect-chain-items pathEnts nil)
        items          (car collected)
        skippedClosed  (cadr collected)
        skippedInvalid (caddr collected)
        chains         (obp:build-chains items)
        placed         0
        processed      (length items)
        skippedCopy    0)
  (foreach chain chains
    (setq result (obp:place-on-chain chain sampleEnts base obp:*spacing* startRef))
    (if (= (cadr result) 'ok)
      (setq placed (+ placed (car result)))
      (setq skippedCopy (1+ skippedCopy))
    )
  )
  (list placed processed skippedClosed skippedInvalid skippedCopy (length chains))
)

(defun obp:place-centerline-curves (pathEnts sampleEnts base pairTol startRef / collected items chains pairedData pairs unpaired pair chain result
                                             placed processed skippedClosed skippedInvalid skippedCopy selfCount stillUnpaired)
  (setq collected      (obp:collect-chain-items pathEnts T)
        items          (car collected)
        skippedClosed  (cadr collected)
        skippedInvalid (caddr collected)
        chains         (obp:build-chains items)
        pairedData     (obp:pair-centerline-chains chains pairTol)
        pairs          (car pairedData)
        unpaired       (cadr pairedData)
        placed         0
        processed      (* 2 (length pairs))
        skippedCopy    0
        selfCount      0
        stillUnpaired  0)
  (foreach pair pairs
    (setq result (obp:place-on-centerline-pair pair sampleEnts base obp:*spacing* startRef))
    (if (= (cadr result) 'ok)
      (setq placed (+ placed (car result)))
      (setq skippedCopy (1+ skippedCopy))
    )
  )
  (foreach chain unpaired
    (if (obp:self-centerline-chain-p chain pairTol)
      (progn
        (setq result (obp:place-on-self-centerline chain sampleEnts base obp:*spacing* startRef)
              selfCount (1+ selfCount)
              processed (1+ processed))
        (if (= (cadr result) 'ok)
          (setq placed (+ placed (car result)))
          (setq skippedCopy (1+ skippedCopy))
        )
      )
      (setq stillUnpaired (1+ stillUnpaired))
    )
  )
  (list placed processed skippedClosed skippedInvalid skippedCopy (length pairs) stillUnpaired selfCount)
)

(defun obp:start-undo (/ doc ret)
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)))
  (setq ret (vl-catch-all-apply 'vla-StartUndoMark (list doc)))
  (not (vl-catch-all-error-p ret))
)

(defun obp:end-undo (/ doc ret)
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)))
  (setq ret (vl-catch-all-apply 'vla-EndUndoMark (list doc)))
  (not (vl-catch-all-error-p ret))
)

(defun obp:run (commandName / oldErr oldEcho undoStarted sampleInfo sampleEnts base targetSpecs
                            areaSS filterResult pathEnts rejected nonCurve placeMode pairTol startRef placeResult
                            placed processed skippedClosed skippedInvalid skippedCopy chainCount unpairedCount selfCount)
  (vl-load-com)
  (setq oldErr *error*
        oldEcho (getvar "CMDECHO")
        undoStarted nil)

  (defun *error* (msg)
    (if undoStarted
      (progn
        (obp:end-undo)
        (setq undoStarted nil)
      )
    )
    (setvar "CMDECHO" oldEcho)
    (if oldErr
      (setq *error* oldErr)
      (setq *error* nil)
    )
    (cond
      ((or (= msg "Function cancelled")
           (= msg "quit / exit abort"))
        (prompt (strcat "\n" commandName " 已取消。"))
      )
      ((and msg (/= msg ""))
        (prompt (strcat "\n錯誤: " msg))
      )
    )
    (princ)
  )

  (setvar "CMDECHO" 0)
  (prompt (strcat "\n=== " commandName " 出線盒自動放置 ==="))

  (setq sampleInfo (obp:get-sample-selection))
  (if (null sampleInfo)
    (prompt "\n未選取有效的出線盒圖例，已取消。")
    (progn
      (setq sampleEnts (car sampleInfo)
            base       (cadr sampleInfo)
            targetSpecs (obp:get-target-legend-specs))
      (if (null targetSpecs)
        (prompt "\n未選取適用的線/曲線圖例，已取消。")
        (progn
          (setq placeMode (obp:get-place-mode))
          (if (= placeMode 'center)
            (setq pairTol (obp:get-center-pair-tol))
            (setq pairTol obp:*center-pair-tol*)
          )
          (setq startRef (obp:get-start-reference))
          (setq areaSS (obp:get-area-selection))
          (if (null areaSS)
            (prompt "\n未選取要放置的範圍，已取消。")
            (progn
              (setq filterResult (obp:filter-area-by-specs areaSS targetSpecs)
                    pathEnts     (car filterResult)
                    rejected     (cadr filterResult)
                    nonCurve     (caddr filterResult))

              (prompt
                (strcat
                  "\n6/6 已過濾可放置物件: "
                  (itoa (length pathEnts))
                  " 個；過濾掉不同圖例曲線 "
                  (itoa rejected)
                  " 個；忽略非線/曲線物件 "
                  (itoa nonCurve)
                  " 個。"
                )
              )

              (if (null pathEnts)
                (prompt "\n範圍內沒有符合適用圖例的物件，已取消。")
                (progn
                  (setq undoStarted (obp:start-undo))

                  (cond
                    ((= placeMode 'center)
                      (prompt "\n雙線中線模式開始放置出線盒（可配對左右邊，或用單一完整物件前半/後半取中線；間距 5000 圖面單位）。")
                    )
                    ((= placeMode 'connect)
                      (prompt "\n相連累計模式開始放置出線盒（相連物件視為同一路徑，間距 5000 圖面單位）。")
                    )
                    (T
                      (prompt "\n一般模式開始放置出線盒（每條物件各自從起點算起，間距 5000 圖面單位）。")
                    )
                  )

                  (setq placeResult
                        (cond
                          ((= placeMode 'center)
                            (obp:place-centerline-curves pathEnts sampleEnts base pairTol startRef)
                          )
                          ((= placeMode 'connect)
                            (obp:place-chained-curves pathEnts sampleEnts base startRef)
                          )
                          (T
                            (obp:place-independent-curves pathEnts sampleEnts base startRef)
                          )
                        )
                        placed         (nth 0 placeResult)
                        processed      (nth 1 placeResult)
                        skippedClosed  (nth 2 placeResult)
                        skippedInvalid (nth 3 placeResult)
                        skippedCopy    (nth 4 placeResult)
                        chainCount     (nth 5 placeResult)
                        unpairedCount  (if (nth 6 placeResult) (nth 6 placeResult) 0)
                        selfCount      (if (nth 7 placeResult) (nth 7 placeResult) 0))

                  (if undoStarted
                    (progn
                      (obp:end-undo)
                      (setq undoStarted nil)
                    )
                  )

                  (prompt
                    (strcat
                      "\n完成，共放置 "
                      (itoa placed)
                      " 個出線盒；已處理物件 "
                      (itoa processed)
                      " 個；略過閉合/無端點物件 "
                      (itoa skippedClosed)
                      " 個；略過無效物件 "
                      (itoa skippedInvalid)
                      " 個；複製失敗 "
                      (itoa skippedCopy)
                      " 個"
                      (cond
                        ((= placeMode 'center)
                          (strcat
                            "；雙線中線配對 "
                            (itoa chainCount)
                            " 組；單一完整物件中線 "
                            (itoa selfCount)
                            " 組；未配對路徑 "
                            (itoa unpairedCount)
                            " 組。"
                          )
                        )
                        ((= placeMode 'connect)
                          (strcat "；相連累計路徑 " (itoa chainCount) " 組。")
                        )
                        (T
                          "。"
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
  )

  (setvar "CMDECHO" oldEcho)
  (if oldErr
    (setq *error* oldErr)
    (setq *error* nil)
  )
  (princ)
)

(defun c:AUTOOUTLETBOX ()
  (obp:run "AUTOOUTLETBOX")
)

(defun c:AUTOOB ()
  (obp:run "AUTOOB")
)

(prompt "\nAUTOOUTLETBOX/AUTOOB 載入完成。")
(princ)
