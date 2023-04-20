# todo
- [ ] test on a machine that does not have MS Excel installed
- [ ] write a navigation map of tasks that the target user might want to have available
  - qform
    - window
      - enum
        - text: What do?
        - symbols
          - [ ] New Workbook
          - [x] Copy Existing Workbook
            - window
              - table
                - text: Workbooks in ``$($PsScriptRoot)``
                - rows
                  ```powershell
                  $(
                    dir "$PsScriptRoot/*.xls*" -Recurse `
                      | select Name, Directory, LastWriteTime `
                  )
                  ```
              - field
                - text: For each worksheet, replace
              - field
                - text: in the name, with
