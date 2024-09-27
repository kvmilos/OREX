# Initial YAML content for the new recording
initial_content = """
name: odks
description: Test recording
start:
  profile: BUSINESS MANAGER
steps:
"""

# Steps template for looping
repeated_steps = """
  - type: invoke
    target:
      - page: Bank Account Ledger Entries
        runtimeRef: {runtime_ref}
      - action: Reverse Transaction
    parameters: {{}}
    description: Invoke <caption>Wycofaj transakcję...</caption>
  - type: page-shown
    source:
      page: Reverse Transaction Entries
    modal: true
    runtimeId: {next_runtime_ref}
    description: Page <caption>Wycofywanie zapisów transakcji - Konto K/G 130-04 Santander Commision 5898 PL</caption> was shown.
  - type: invoke
    target:
      - page: Reverse Transaction Entries
        runtimeRef: {next_runtime_ref}
      - action: Reverse
    parameters: {{}}
    description: Invoke <caption>Wycofaj</caption>
  - type: page-shown
    source:
      page: null
      automationId: 8da61efd-0002-0003-0507-0b0d1113171d
      caption: Potwierdź
    modal: true
    runtimeId: {confirm_runtime_ref}
    description: Page <caption>Potwierdź</caption> was shown.
  - type: invoke
    target:
      - page: null
        automationId: 8da61efd-0002-0003-0507-0b0d1113171d
        caption: Potwierdź
        runtimeRef: {confirm_runtime_ref}
    invokeType: Yes
    description: Invoke Yes on <caption>Potwierdź</caption>
  - type: page-closed
    source:
      page: null
    runtimeId: {confirm_runtime_ref}
    description: Page <caption>Potwierdź</caption> was closed.
  - type: page-closed
    source:
      page: Reverse Transaction Entries
    runtimeId: {next_runtime_ref}
    description: Page <caption>Wycofywanie zapisów transakcji - Konto K/G 130-04 Santander Commision 5898 PL</caption> was closed.
  - type: message
    automationId: 8da61efd-0002-0003-0507-0b0d1113171d
    text: Zapisy zostały pomyślnie wycofane.
    description: "Message: <value>Zapisy zostały pomyślnie wycofane.</value>"
"""

# Initialize the YAML content
yaml_content = initial_content

# Loop to generate the repeated steps for each iteration
for i in range(100):  # Adjust the range as needed
    runtime_ref = f"b9{i}bp"
    next_runtime_ref = f"b9{i}jg"
    confirm_runtime_ref = f"b9{i}of"
    
    yaml_content += repeated_steps.format(
        runtime_ref=runtime_ref,
        next_runtime_ref=next_runtime_ref,
        confirm_runtime_ref=confirm_runtime_ref
    )

# Save the generated YAML content to a file
with open("odks_recording.yaml", "w") as f:
    f.write(yaml_content)
