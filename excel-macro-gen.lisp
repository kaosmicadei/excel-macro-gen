(defstruct (cell (:constructor cell (column row)))
  column
  row)

(defmethod print-object ((object cell) stream)
  (with-slots (column row) object
    (format stream "~a~a" column row)))

(defstruct range
  first-cell
  last-cell)

(defun range (first-column first-row last-column last-row)
  (make-range :first-cell (cell first-column first-row)
              :last-cell  (cell last-column last-row)))

(defmethod print-object ((object range) stream)
  (with-slots (first-cell last-cell) object
    (format stream "~a:~a" first-cell last-cell)))

;; simplify the access to range fields
(defmacro with-range-fields ((a b c d) range &body body)
  `(let ((,a (cell-column (range-first-cell ,range)))
         (,b (cell-row   (range-first-cell ,range)))
         (,c (cell-column (range-last-cell  ,range)))
         (,d (cell-row   (range-last-cell  ,range))))
        ,@body))

(defun expand-range (range)
  (with-range-fields (fstcol fstrow lstcol lstrow) range
    (loop for col from (char-code fstcol) to (char-code lstcol)
      append (loop for row from fstrow to lstrow
               collect (cell (code-char col) row)))))

(defun expand-columns (range)
  (with-range-fields (fstcol fstrow lstcol lstrow) range
    (loop for col from (char-code fstcol) to (char-code lstcol)
      collect (range (code-char col) fstrow (code-char col) lstrow))))

(defun expand-rows (range)
  (with-range-fields (fstcol fstrow lstcol lstrow) range
    (loop for row from fstrow to lstrow
      collect (range fstcol row lstcol row))))

(defun each-column (range)
  (with-range-fields (fstcol fstrow lstcol lstrow) range
    (declare (ignorable fstcol))
    (declare (ignorable lstcol))
    (if (equal fstrow lstrow)
      (expand-range   range)
      (expand-columns range))))

(defun each-row (range)
  (with-range-fields (fstcol fstrow lstcol lstrow) range
    (declare (ignorable fstrow))
    (declare (ignorable lstrow))
    (if (equal fstcol lstcol)
      (expand-range range)
      (expand-rows  range))))

(defstruct (build-block (:constructor build-block (function-name arguments)))
  function-name
  arguments)

(defmethod print-object ((object build-block) stream)
  (with-slots (function-name arguments) object
    (format stream "~:@(~a~)(~{~:@(~a~)~^;~})" function-name arguments)))

(defstruct (binary-operator (:constructor binary-operator (operator left right)))
  operator
  left
  right)

(defmethod print-object ((object binary-operator) stream)
  (with-slots (operator left right) object
    (format stream "~a ~a ~a" left operator right)))

(defstruct (formula (:constructor formula (name output expression)))
  name
  output
  expression)

(defmethod print-object ((object formula) stream)
  (with-slots (output expression) object
    (if (and (typep output 'RANGE) (listp expression))
      (let ((formulas (mapcar (lambda (o e) (list o e)) (expand-range output) expression)))
        (format stream "~{~{Range(\"~a\").Formula = \"~a\"~}~%~}" formulas))
      (format stream "Range(\"~a\").Formula = \"~a\"" output expression))))

(defgeneric block-sum (object)
  (:documentation "Excel sum of a single range or a list of sum for multiple ranges"))

(defmethod block-sum ((object range))
  (build-block "sum" (list object)))

(defmethod block-sum ((object list))
  (loop for range in object
    collect (block-sum range)))

(defgeneric block-add (left right)
  (:documentation "Add to expression"))

(defmethod block-add ((left build-block) (right build-block))
  (binary-operator "+" left right))

(defmethod block-add ((left build-block) (right list))
  (loop for rhs in right
    collect (binary-operator "+" left rhs)))

(defmethod block-add ((left list) (right build-block))
  (loop for lhs in left
    collect (binary-operator "+" lhs right)))

(defmethod block-add ((left list) (right list))
  (mapcar (lambda (lhs rhs) (binary-operator "+" lhs rhs)) left right))

;; AREA 51
(defvar *range1* (range #\A 1 #\B 1))
(defvar *range2* (range #\A 1 #\B 3))

(defvar *formula1*
  (formula
    "myformula"
    (cell #\A 5)
    (block-sum
      (range #\A 1 #\A 3))))

(defvar *formula2*
  (formula
    "myformula"
    (range #\A 5 #\B 5)
    (block-sum
      (each-column (range #\A 1 #\B 3)))))

(defvar *formula3*
  (formula
    "myformula"
    (range #\C 1 #\C 10)
    (block-add
      (each-row (range #\A 1 #\A 10))
      (each-row (range #\B 1 #\B 10)))))

(defvar *formula4*
  (formula
    "myformula"
    (range #\H 1 #\I 10)
    (block-add
      (block-sum (each-row (range #\A 1 #\B 10)))
      (block-sum (each-row (range #\C 1 #\D 10))))))

(print *formula3*)

(quit)
