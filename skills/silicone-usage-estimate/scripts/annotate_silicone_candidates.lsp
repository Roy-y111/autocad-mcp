(vl-load-com)

(defun sca-outdir ( / d)
  (setq d (getenv "SILICONE_OUT_DIR"))
  (if (and d (/= d ""))
    (strcat d "/")
    (strcat (getenv "USERPROFILE") "/Desktop/CL_CAD/outputs/")))

(defun sca-split (s d / p out n)
  (setq out '() n (strlen d))
  (while (setq p (vl-string-search d s))
    (setq out (append out (list (substr s 1 p))) s (substr s (+ p n 1))))
  (append out (list s)))

(defun sca-erase (ss / i)
  (if ss (progn (setq i 0) (repeat (sslength ss) (entdel (ssname ss i)) (setq i (1+ i))))))

(defun sca-near-dimstyle (x0 y0 x1 y1 w h / ss sty)
  (setq ss (ssget "C" (list (- x0 (* w 1.5)) (- y0 (* h 1.5)))
                         (list (+ x1 (* w 1.5)) (+ y1 (* h 1.5))) '((0 . "DIMENSION"))))
  (if ss
    (progn
      (setq sty (cdr (assoc 3 (entget (ssname ss 0)))))
      (if sty (vl-catch-all-apply 'vl-cmdf (list "_.-DIMSTYLE" "_R" sty))))))

(defun sca-finalize (/ txt obj ins)
  (if sca-code
    (progn
      (setq txt (strcat sca-code " 玻璃矽利康"
                        sca-lines
                        "\P單面合計=" (rtos sca-total 2 1)
                        "\P面數=2"
                        "\P兩面總長=" (rtos (* sca-total 2.0) 2 1)))
      (setq ins (list sca-minx (- sca-miny 35.0) 0.0))
      (setq obj (vla-AddMText sca-ms (vlax-3d-point ins) 280.0 txt))
      (vla-put-Layer obj "SILICONE-ANNOTATION")
      (vla-put-Color obj 2)
      (vla-put-Height obj 10.0))))

(defun c:SILANNOTATE (/ doc lays lay f line p code pane x y w h x0 x1 y0 y1 perim obj oldlay oldos olddimlayer idx)
  (setq oldlay (getvar "CLAYER") oldos (getvar "OSMODE") olddimlayer (getvar "DIMLAYER"))
  (setvar "OSMODE" 0)
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)) sca-ms (vla-get-ModelSpace doc) lays (vla-get-Layers doc))
  (setq lay (vl-catch-all-apply 'vla-Item (list lays "SILICONE-ANNOTATION")))
  (if (vl-catch-all-error-p lay) (setq lay (vla-Add lays "SILICONE-ANNOTATION")))
  (vla-put-Color lay 2)
  (sca-erase (ssget "_X" '((8 . "SILICONE-ANNOTATION"))))
  (setvar "CLAYER" "SILICONE-ANNOTATION")
  (setvar "DIMLAYER" "SILICONE-ANNOTATION")
  (setq f (open (strcat (sca-outdir) "silicone_seeds.txt") "r"))
  (setq sca-code nil sca-lines "" sca-total 0.0 sca-minx 0.0 sca-miny 0.0 idx 0)
  (while (setq line (read-line f))
    (setq p (sca-split line "|") code (nth 0 p) pane (nth 1 p)
          x (atof (nth 2 p)) y (atof (nth 3 p)) w (atof (nth 4 p)) h (atof (nth 5 p)))
    (if (or (not sca-code) (/= code sca-code))
      (progn
        (sca-finalize)
        (setq sca-code code sca-lines "" sca-total 0.0 sca-minx 1e99 sca-miny 1e99 idx 0)))
    (setq idx (1+ idx) x0 (- x (/ w 2.0)) x1 (+ x (/ w 2.0)) y0 (- y (/ h 2.0)) y1 (+ y (/ h 2.0)))
    (setq sca-minx (min sca-minx x0) sca-miny (min sca-miny y0))
    (sca-near-dimstyle x0 y0 x1 y1 w h)
    ;; Same dimension method as SiliconeCalc.lsp.
    (command "_.DIMLINEAR" "_NON" (list x0 y1) "_NON" (list x1 y1) "_NON" (list x (+ y1 (max 12.0 (* h 0.10)))))
    (setq obj (vlax-ename->vla-object (entlast))) (vla-put-Layer obj "SILICONE-ANNOTATION") (vla-put-Color obj 2)
    (command "_.DIMLINEAR" "_NON" (list x0 y0) "_NON" (list x0 y1) "_NON" (list (- x0 (max 12.0 (* w 0.10))) y))
    (setq obj (vlax-ename->vla-object (entlast))) (vla-put-Layer obj "SILICONE-ANNOTATION") (vla-put-Color obj 2)
    (setq perim (* 2.0 (+ w h)) sca-total (+ sca-total perim))
    (setq sca-lines (strcat sca-lines "\P" (itoa idx) "." pane "  (" (rtos w 2 1) "+" (rtos h 2 1) ")x2=" (rtos perim 2 1))))
  (sca-finalize)
  (close f)
  (setvar "CLAYER" oldlay) (setvar "OSMODE" oldos) (setvar "DIMLAYER" olddimlayer)
  (princ "\nSILANNOTATE complete: dimensions and formulas created for all candidates.")
  (princ))

