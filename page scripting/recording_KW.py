# Initial YAML content for the new recording
initial_content = """
name: KW Recording
description: Test recording for KW
start:
  profile: BUSINESS MANAGER
steps:
"""

# Steps template for looping
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

# Loop to generate the repeated steps for each iteration
for i in range(100):  # Adjust the range as needed
    runtime_ref = f"bw{i+1}"
    next_runtime_ref = f"bx{i+5}"
    
    yaml_content += repeated_steps.format(
        runtime_ref=runtime_ref,
        next_runtime_ref=next_runtime_ref
    )

# Save the generated YAML content to a file
with open("KW_recording.yaml", "w") as f:
    f.write(yaml_content)

# Initial YAML content for the new recording
initial_content = """
name: KW2 Recording
description: Test recording for KW2
start:
  profile: BUSINESS MANAGER
steps:
"""

# Steps template for looping
repeated_steps = """
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

# Initialize the YAML content
yaml_content = initial_content

# Loop to generate the repeated steps for each iteration
for i in range(300):  # Adjust the range as needed
    runtime_ref = f"bap{i+1}"
    next_runtime_ref = f"bap{i+2}"
    
    yaml_content += repeated_steps.format(
        runtime_ref=runtime_ref,
        next_runtime_ref=next_runtime_ref
    )

# Save the generated YAML content to a file
with open("KW2_recording.yaml", "w") as f:
    f.write(yaml_content)