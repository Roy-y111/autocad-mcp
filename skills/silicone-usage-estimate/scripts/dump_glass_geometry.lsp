(vl-load-com)

(defun sg-outdir ( / d)
  (setq d (getenv "SILICONE_OUT_DIR"))
  (if (and d (/= d ""))
    (strcat d "/")
    (strcat (getenv "USERPROFILE") "/Desktop/CL_CAD/outputs/")))

(defun sg-error (msg / ef)
  (setq ef (open (strcat (sg-outdir) "a6_glass_error.txt") "w"))
  (if ef (progn
    (write-line (strcat "index=" (itoa sg-index) " object=" sg-oname " stage=" sg-stage) ef)
    (write-line msg ef)
    (close ef)))
  (princ))

(defun sg-val (obj prop / result)
  (setq result (vl-catch-all-apply 'vlax-get-property (list obj prop)))
  (if (vl-catch-all-error-p result) nil result))

(defun sg-pt-list (value)
  (cond
    ((= (type value) 'VARIANT) (sg-pt-list (vlax-variant-value value)))
    ((= (type value) 'SAFEARRAY) (vlax-safearray->list value))
    ((listp value) value)
    (t nil)))

(defun sg-num (n) (rtos n 2 6))

(defun sg-write-coords (f coords / text)
  (setq text "")
  (foreach n coords (setq text (strcat text (if (= text "") "" ",") (sg-num n))))
  (write-line text f))

(defun c:SGDUMPGLASS (/ *error* doc ms f obj oname layer bname pt attrs att tag val coords closed)
  (setq *error* 'sg-error)
  (setq sg-index 0 sg-oname "START" sg-stage "START")
  (setq f (open (strcat (sg-outdir) "a6_glass_raw.txt") "w"))
  (if (not f)
    (princ "\nSGDUMPGLASS: cannot open output file.")
    (progn
      (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)))
      (setq ms (vla-get-ModelSpace doc))
      (vlax-for obj ms
        (setq sg-index (1+ sg-index))
        (setq oname (vla-get-ObjectName obj))
        (setq sg-oname oname)
        (setq layer (vla-get-Layer obj))
        (cond
          ((= oname "AcDbBlockReference")
            (setq sg-stage "block-name")
            (setq bname (sg-val obj 'EffectiveName))
            (if (not bname) (setq bname (vla-get-Name obj)))
            (setq sg-stage (strcat "block-member-" bname))
            (if (member bname '("_LCNo001" "_LDNO001" "_LOPN001"))
              (progn
                (setq sg-stage "block-point")
                (setq pt (sg-pt-list (vla-get-InsertionPoint obj)))
                (setq sg-stage "block-hasattrs")
                (if (= :vlax-true (vla-get-HasAttributes obj))
                  (progn
                    (setq sg-stage "block-attrs")
                    (setq attrs (vlax-invoke obj 'GetAttributes))
                    (foreach att attrs
                      (setq tag (strcase (vla-get-TagString att)))
                      (if (member tag '("NO" "U"))
                        (progn
                          (setq val (vla-get-TextString att))
                          (write-line
                            (strcat "TAG|" bname "|" val "|" (sg-num (car pt)) "|" (sg-num (cadr pt))) f)))))))))
          ((= oname "AcDbAttributeDefinition")
            (setq tag (sg-val obj 'TagString))
            (if (and tag (= (strcase tag) "FIX"))
              (progn
                (setq pt (sg-pt-list (vla-get-InsertionPoint obj)))
                (write-line (strcat "FIX|" (sg-num (car pt)) "|" (sg-num (cadr pt))) f))))
          ((and (= oname "AcDbPolyline") (member layer '("OPEN" "C7-0")))
            (setq coords (sg-pt-list (vla-get-Coordinates obj)))
            (setq closed (if (= :vlax-true (vla-get-Closed obj)) "1" "0"))
            (write-line (strcat "PL|" layer "|" closed "|") f)
            (sg-write-coords f coords))))
      (close f)
      (princ (strcat "\nSGDUMPGLASS complete: " (sg-outdir) "a6_glass_raw.txt"))))
  (princ))
