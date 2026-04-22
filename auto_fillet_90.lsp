(vl-load-com)

(defun af:v+ (a b)
  (list (+ (car a) (car b)) (+ (cadr a) (cadr b)))
)

(defun af:v- (a b)
  (list (- (car a) (car b)) (- (cadr a) (cadr b)))
)

(defun af:v* (v s)
  (list (* (car v) s) (* (cadr v) s))
)

(defun af:dot (a b)
  (+ (* (car a) (car b)) (* (cadr a) (cadr b)))
)

(defun af:crossz (a b)
  (- (* (car a) (cadr b)) (* (cadr a) (car b)))
)

(defun af:len (v)
  (sqrt (af:dot v v))
)

(defun af:pt2d (p)
  (list (car p) (cadr p))
)

(defun af:take (lst n / out)
  (setq out '())
  (while (and lst (> n 0))
    (setq out (cons (car lst) out)
          lst (cdr lst)
          n   (1- n))
  )
  (reverse out)
)

(defun af:drop (lst n)
  (while (and lst (> n 0))
    (setq lst (cdr lst)
          n   (1- n))
  )
  lst
)

(defun af:set-dxf (code val el / old)
  (if (setq old (assoc code el))
    (subst (cons code val) old el)
    (append el (list (cons code val)))
  )
)

(defun af:append-dxf-if (lst code el)
  (if (assoc code el)
    (append lst (list (assoc code el)))
    lst
  )
)

(defun af:split-geom (el / header tail inGeom d c)
  (setq header '()
        tail   '()
        inGeom nil)

  (foreach d el
    (setq c (car d))
    (cond
      ((and (not inGeom) (= c 10))
       (setq inGeom T)
      )
      ((and inGeom (member c '(10 40 41 42 91)))
       ; skip vertex groups
      )
      (inGeom
       (setq tail (cons d tail))
      )
      (T
       (setq header (cons d header))
      )
    )
  )

  (list (reverse header) (reverse tail))
)

(defun af:parse-nodes (el / nodes d)
  ; node format: (point2d bulgeToNext)
  (setq nodes '())
  (foreach d el
    (cond
      ((= 10 (car d))
       (setq nodes (cons (list (af:pt2d (cdr d)) 0.0) nodes))
      )
      ((and (= 42 (car d)) nodes)
       (setq nodes (cons (list (caar nodes) (cdr d)) (cdr nodes)))
      )
    )
  )
  (reverse nodes)
)

(defun af:parse-heavy-nodes (ent / e ed typ nodes p b)
  ; parse classic 2D POLYLINE vertices
  (setq nodes '()
        e     (entnext ent))
  (while e
    (setq ed  (entget e)
          typ (cdr (assoc 0 ed)))
    (cond
      ((= typ "SEQEND")
       (setq e nil)
      )
      ((= typ "VERTEX")
       (setq p (af:pt2d (cdr (assoc 10 ed)))
             b (cdr (assoc 42 ed)))
       (if (null b)
         (setq b 0.0)
       )
       (setq nodes (append nodes (list (list p b))))
      )
    )
    (if e
      (setq e (entnext e))
    )
  )
  nodes
)

(defun af:is-right-angle (v1 v2 tol / l1 l2)
  (setq l1 (af:len v1)
        l2 (af:len v2))
  (if (or (<= l1 1e-9) (<= l2 1e-9))
    nil
    (<= (abs (/ (af:dot v1 v2) (* l1 l2))) tol)
  )
)

(defun af:pick-radius (len1 len2 radii eps / out r)
  (setq out nil)
  (while (and radii (null out))
    (setq r (car radii))
    (if (and (>= (- len1 r) eps)
             (>= (- len2 r) eps))
      (setq out r)
    )
    (setq radii (cdr radii))
  )
  out
)

(defun af:set-last-bulge-zero (nodes / rev lastNode)
  (if nodes
    (progn
      (setq rev (reverse nodes)
            lastNode (car rev))
      (reverse (cons (list (car lastNode) 0.0) (cdr rev)))
    )
    nodes
  )
)

(defun af:find-new-polyline (before / e typ out)
  ; scan all newly created entities after BEFORE, keep the newest polyline
  (setq e (if before (entnext before) (entnext))
        out nil)
  (while e
    (setq typ (cdr (assoc 0 (entget e))))
    (if (or (= typ "LWPOLYLINE") (= typ "POLYLINE"))
      (setq out e)
    )
    (setq e (entnext e))
  )
  out
)

(defun af:replace-with-lwpolyline (ent nodes srcFlag / el lwFlag data nd newEnt)
  ; rebuild as LWPOLYLINE (used for classic POLYLINE)
  (setq el     (entget ent)
        lwFlag (if (/= 0 (logand srcFlag 128)) 128 0)
        data   (list '(0 . "LWPOLYLINE")
                      '(100 . "AcDbEntity")))

  (setq data (af:append-dxf-if data 67 el))
  (setq data (af:append-dxf-if data 410 el))
  (setq data (af:append-dxf-if data 8 el))
  (setq data (af:append-dxf-if data 6 el))
  (setq data (af:append-dxf-if data 48 el))
  (setq data (af:append-dxf-if data 60 el))
  (setq data (af:append-dxf-if data 62 el))
  (setq data (af:append-dxf-if data 370 el))
  (setq data (af:append-dxf-if data 420 el))
  (setq data (af:append-dxf-if data 430 el))
  (setq data (af:append-dxf-if data 440 el))

  (setq data (append data
                     (list '(100 . "AcDbPolyline")
                           (cons 90 (length nodes))
                           (cons 70 lwFlag))))

  (setq data (af:append-dxf-if data 38 el))
  (setq data (af:append-dxf-if data 39 el))
  (setq data (af:append-dxf-if data 210 el))

  (foreach nd nodes
    (setq data (append data (list (cons 10 (car nd)))))
    (if (> (abs (cadr nd)) 1e-12)
      (setq data (append data (list (cons 42 (cadr nd)))))
    )
  )

  (setq newEnt (entmakex data))
  (if newEnt
    (progn
      (entdel ent)
      newEnt
    )
    nil
  )
)

(defun af:fillet-90-on-entity (ent / el typ flag nodes i cnt prev curr nxt
                                p0 p1 p2 bPrev bCurr v1 v2 l1 l2 cross
                                rad u1 u2 pA pB arcBulge sp header tail newGeom newEl nd)
  (cond
    ((null ent)
     (prompt "\n未指定圖元。")
    )

    (T
      (setq el   (entget ent)
            typ  (cdr (assoc 0 el))
            flag (cdr (assoc 70 el)))

      (cond
        ((not (member typ '("LWPOLYLINE" "POLYLINE")))
          (prompt "\n只支援 LWPOLYLINE / POLYLINE。")
        )

        ((and (= typ "POLYLINE")
              (/= 0 (logand flag (+ 8 16 64))))
          (prompt "\n僅支援 2D POLYLINE。")
        )

        ((/= 0 (logand 1 flag))
          (prompt "\n目前僅支援未封閉(open)的 polyline。")
        )

        (T
          (setq nodes (if (= typ "LWPOLYLINE")
                        (af:parse-nodes el)
                        (af:parse-heavy-nodes ent)))

          (if (< (length nodes) 3)
            (prompt "\n頂點不足，至少需要 3 個頂點。")
            (progn
              (setq i 1
                    cnt 0)

              ; 從起點往後依序處理內角
              (while (< i (1- (length nodes)))
                (setq prev  (nth (1- i) nodes)
                      curr  (nth i nodes)
                      nxt   (nth (1+ i) nodes)
                      p0    (car prev)
                      p1    (car curr)
                      p2    (car nxt)
                      bPrev (cadr prev)
                      bCurr (cadr curr))

                (if (and (< (abs bPrev) 1e-9)
                         (< (abs bCurr) 1e-9))
                  (progn
                    (setq v1 (af:v- p1 p0)
                          v2 (af:v- p2 p1)
                          l1 (af:len v1)
                          l2 (af:len v2)
                          cross (af:crossz v1 v2))

                    (if (and (> (abs cross) 1e-9)
                             (af:is-right-angle v1 v2 1e-3))
                      (progn
                        (setq rad (af:pick-radius l1 l2 '(150.0 100.0 50.0) 1e-6))
                        (if rad
                          (progn
                            (setq u1 (af:v* v1 (/ 1.0 l1))
                                  u2 (af:v* v2 (/ 1.0 l2))
                                  pA (af:v- p1 (af:v* u1 rad))
                                  pB (af:v+ p1 (af:v* u2 rad))
                                  arcBulge (* (if (> cross 0.0) 1.0 -1.0)
                                              0.4142135623730951)
                                  nodes (append (af:take nodes i)
                                                (list (list pA arcBulge)
                                                      (list pB 0.0))
                                                (af:drop nodes (1+ i)))
                                  i    (+ i 2)
                                  cnt  (1+ cnt))
                          )
                          (setq i (1+ i))
                        )
                      )
                      (setq i (1+ i))
                    )
                  )
                  (setq i (1+ i))
                )
              )

              (setq nodes (af:set-last-bulge-zero nodes))

              (if (= typ "LWPOLYLINE")
                (progn
                  (setq sp     (af:split-geom el)
                        header (af:set-dxf 90 (length nodes) (car sp))
                        tail   (cadr sp)
                        newGeom '())

                  (foreach nd nodes
                    (setq newGeom (append newGeom (list (cons 10 (car nd)))))
                    (if (> (abs (cadr nd)) 1e-12)
                      (setq newGeom (append newGeom (list (cons 42 (cadr nd)))))
                    )
                  )

                  (setq newEl (append header newGeom tail))

                  (if (entmod newEl)
                    (progn
                      (entupd ent)
                      (prompt (strcat "\n完成，已處理 " (itoa cnt) " 個 90 度轉角。"))
                      cnt
                    )
                    (prompt "\n更新失敗，請檢查圖元狀態。")
                  )
                )
                (if (af:replace-with-lwpolyline ent nodes flag)
                  (progn
                    (prompt (strcat "\n完成，已處理 " (itoa cnt) " 個 90 度轉角。"))
                    cnt
                  )
                  (prompt "\n更新失敗，無法重建為 LWPOLYLINE。")
                )
              )
            )
          )
        )
      )
    )
  )
)

(defun c:PL90F (/ ss i ent total)
  (prompt "\n框選/多選要自動圓角的 PLINE (僅 LWPOLYLINE/POLYLINE): ")
  (setq ss (ssget '((0 . "LWPOLYLINE,POLYLINE"))))
  (if ss
    (progn
      (setq i 0
            total (sslength ss))
      (while (< i total)
        (setq ent (ssname ss i))
        (af:fillet-90-on-entity ent)
        (setq i (1+ i))
      )
      (prompt (strcat "\nPL90F 完成，共處理 " (itoa total) " 條 PLINE。"))
    )
    (prompt "\n已取消或未選到 PLINE。")
  )

  (princ)
)

(defun af:draw-pline-and-wait ()
  ; keep yielding PAUSE until PLINE is fully finished by user
  (command "_.PLINE")
  (while (/= 0 (logand 1 (getvar "CMDACTIVE")))
    ; enforce ORTHO during PL90FA drawing
    (if (/= (getvar "ORTHOMODE") 1)
      (setvar "ORTHOMODE" 1)
    )
    (command pause)
  )
)

(defun c:PL90FA (/ before pl oldOrtho oldErr)
  (setq before   (entlast)
        oldOrtho (getvar "ORTHOMODE")
        oldErr   *error*)

  (defun *error* (msg)
    (setvar "ORTHOMODE" oldOrtho)
    (if oldErr
      (setq *error* oldErr)
      (setq *error* nil)
    )
    (if (and msg
             (/= msg "Function cancelled")
             (/= msg "quit / exit abort"))
      (prompt (strcat "\n錯誤: " msg))
    )
    (princ)
  )

  ; force ortho mode during PL90FA
  (if (/= oldOrtho 1)
    (setvar "ORTHOMODE" 1)
  )

  (prompt "\n開始繪製 PLINE，結束後會自動套用 90 度圓角...")
  (af:draw-pline-and-wait)
  (setq pl (af:find-new-polyline before))

  (cond
    ((null pl)
      (prompt "\n未找到剛建立的 polyline。")
    )
    (T
      (af:fillet-90-on-entity pl)
    )
  )

  (setvar "ORTHOMODE" oldOrtho)
  (if oldErr
    (setq *error* oldErr)
    (setq *error* nil)
  )

  (princ)
)

(prompt "\n載入完成。PL90F=框選/多選 PLINE 後圓角；PL90FA=畫完 PLINE 後自動圓角。")
(princ)
