# Original YAML content
original_content = """
name: Recording
description: Test recording
start:
  profile: BUSINESS MANAGER
steps:

  - type: page-closed
    source:
      page: ITI Purch. Doc. Nos. Tmp. List
    runtimeId: b13j
    description: Page <caption>Lista szablon�w serii num. dok. zakupu</caption> was closed.
  - type: page-shown
    source:
      page: ITI Purch. Doc. Nos. Tmp. List
    modal: true
    runtimeId: b14n
    description: Page <caption>Lista szablon�w serii num. dok. zakupu</caption> was shown.

  - type: set-current-row
    target:
      - page: ITI Purch. Doc. Nos. Tmp. List
        runtimeRef: b14j
      - repeater: General
    targetRecord:
      relative: 4
    description: Set current row in <caption>General</caption>
  - type: invoke
    target:
      - page: ITI Purch. Doc. Nos. Tmp. List
        runtimeRef: b14j
      - repeater: General
    description: Invoke row on <caption>General</caption>
  - type: page-closed
    source:
      page: ITI Purch. Doc. Nos. Tmp. List
    runtimeId: b14j
    description: Page <caption>Lista szablon�w serii num. dok. zakupu</caption> was closed.
"""

# New steps template
new_steps_template = """
  - type: invoke
    target:
      - page: null
        automationId: 8da61efd-0002-0003-0507-0b0d1113171d
        caption: Potwierdź
        runtimeRef: {runtime_ref}
    invokeType: Yes
    description: Invoke Yes on <caption>Potwierdź</caption>
  - type: page-closed
    source:
      page: null
    runtimeId: {runtime_ref}
    description: Page <caption>Potwierdź</caption> was closed.
  - type: page-shown
    source:
      page: null
      automationId: 8da61efd-0002-0003-0507-0b0d1113171d
      caption: Potwierdź
    modal: true
    runtimeId: {next_runtime_ref}
    description: Page <caption>Potwierdź</caption> was shown.
"""

# Initialize the new YAML content
new_yaml_content = """
name: Recording
description: Test recording
start:
  profile: BUSINESS MANAGER
steps:
"""

# Loop to generate the new steps
for i in range(100):  # Adjust the range as needed
    runtime_ref = f"b8c{i}h"
    next_runtime_ref = f"b8c{i+1}n"
    new_yaml_content += new_steps_template.format(runtime_ref=runtime_ref, next_runtime_ref=next_runtime_ref)

# Save the new content to a file
with open("recording2.yaml", "w") as f:
    f.write(new_yaml_content)