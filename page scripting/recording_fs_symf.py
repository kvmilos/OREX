# Initial YAML content for the new recording
initial_content = """
name: FS_symf Recording
description: Test recording for FS_symf
start:
  profile: BUSINESS MANAGER
steps:
"""

# Steps template for looping
repeated_steps = """
  - type: set-current-row
    target:
      - page: ITI Sales Doc. Nos. Tmp. List
        runtimeRef: {runtime_ref}
      - repeater: General
    targetRecord:
      relative: 1
    description: Set current row in <caption>General</caption>
  - type: invoke
    target:
      - page: ITI Sales Doc. Nos. Tmp. List
        runtimeRef: {runtime_ref}
      - repeater: General
    description: Invoke row on <caption>General</caption>
  - type: page-closed
    source:
      page: ITI Sales Doc. Nos. Tmp. List
    runtimeId: {runtime_ref}
    description: Page <caption>Lista szablonów serii num. dok. sprzedaży</caption> was closed.
  - type: page-shown
    source:
      page: ITI Sales Doc. Nos. Tmp. List
    modal: true
    runtimeId: {next_runtime_ref}
    description: Page <caption>Lista szablonów serii num. dok. sprzedaży</caption> was shown.
"""

# Initialize the YAML content
yaml_content = initial_content

# Loop to generate the repeated steps for each iteration
for i in range(100):  # Adjust the range as needed
    runtime_ref = f"b{i+351}"
    next_runtime_ref = f"b{i+365}"
    
    yaml_content += repeated_steps.format(
        runtime_ref=runtime_ref,
        next_runtime_ref=next_runtime_ref
    )

# Save the generated YAML content to a file
with open("FS_symf_recording.yaml", "w") as f:
    f.write(yaml_content)

# Initial YAML content for the new recording
initial_content = """
name: FS_symf2 Recording
description: Test recording for FS_symf2
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
  - type: message
    automationId: 8da61efd-0002-0003-0507-0b0d1113171d
    text: Zmieniono Data księgowania w zamówieniu sprzedaży, co może mieć wpływ na
      ceny i rabaty w wierszach zamówienia sprzedaży. Należy przejrzeć wiersze
      oraz ręcznie zatwierdzić ceny i rabaty w razie konieczności.
    description: "Message: <value>Zmieniono Data księgowania w zamówieniu sprzedaży,
      co może mieć wpływ na ceny i rabaty w wierszach zamówienia sprzedaży.
      Należy przejrzeć wiersze oraz ręcznie zatwierdzić ceny i rabaty w razie
      konieczności.</value>"
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
for i in range(900):  # Adjust the range as needed
    runtime_ref = f"bv{i+1}"
    next_runtime_ref = f"bv{i+2}"
    
    yaml_content += repeated_steps.format(
        runtime_ref=runtime_ref,
        next_runtime_ref=next_runtime_ref
    )

# Save the generated YAML content to a file
with open("FS_symf2_recording.yaml", "w") as f:
    f.write(yaml_content)