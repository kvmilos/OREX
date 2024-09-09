initial_content = """
name: Recording
description: Test recording
start:
  profile: BUSINESS MANAGER
steps:
"""

repeated_steps = """
  - type: set-current-row
    target:
      - page: ITI Purch. Doc. Nos. Tmp. List
        runtimeRef: {runtime_ref}
      - repeater: General
    targetRecord:
      relative: 4
    description: Set current row in <caption>General</caption>
  - type: invoke
    target:
      - page: ITI Purch. Doc. Nos. Tmp. List
        runtimeRef: {runtime_ref}
      - repeater: General
    description: Invoke row on <caption>General</caption>
  - type: page-closed
    source:
      page: ITI Purch. Doc. Nos. Tmp. List
    runtimeId: {runtime_ref}
    description: Page <caption>Lista szablonów serii num. dok. zakupu</caption> was closed.
  - type: page-shown
    source:
      page: ITI Purch. Doc. Nos. Tmp. List
    modal: true
    runtimeId: {next_runtime_ref}
    description: Page <caption>Lista szablonów serii num. dok. zakupu</caption> was shown.
"""

# Initialize the YAML content
yaml_content = initial_content

# Loop to generate the repeated steps
for i in range(100):
    runtime_ref = f"b{i+13}j"
    next_runtime_ref = f"b{i+14}n"
    yaml_content += repeated_steps.format(runtime_ref=runtime_ref, next_runtime_ref=next_runtime_ref)

# save
with open("recording1.yaml", "w") as f:
    f.write(yaml_content)