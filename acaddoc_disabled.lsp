;; Define the SetVars function
(defun SetVars ()
  ;; Set FILEDIA to 1
  (setvar "FILEDIA" 1)
  (princ "\nFILEDIA set to 1")

  ;; Set MIRRTEXT to 0
  (setvar "MIRRTEXT" 0)
  (princ "\nMIRRTEXT set to 0")
)

;; Define the UnloadARX function
(defun UnloadARX ()
  ;; List of known ARX applications that might be loaded
  ;; You may add more known ARX names that you want to unload.
  (setq known-arx-list '("test.arx")) ;; Example ARX names

  ;; Attempt to unload all ARX applications in the list
  (foreach arx known-arx-list
    (if (arxunload arx)
      (princ (strcat "\nUnloaded ARX: " arx))
      (princ (strcat "\nARX not loaded: " arx))
    )
  )
)

;; Automatically execute functions on each load
(SetVars)
(UnloadARX)

;; Notify user that the script has completed
(princ "\nacaddoc.lsp loaded successfully.")
(princ)
