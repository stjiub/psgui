@{
    Severity     = @('Error', 'Warning')
    ExcludeRules = @(
        # WPF event handlers take ($sender, $e) parameters; they are parameters here, not assignments
        'PSAvoidAssignmentToAutomaticVariable'
        # Handlers must match the event delegate signature even when they ignore an argument
        'PSReviewUnusedParameter'
        # Internal UI helpers; -WhatIf/-Confirm would not make sense for them
        'PSUseShouldProcessForStateChangingFunctions'
        # Helper names (Load-DataFile, Toggle-ShellGrid, ...) read better as plain verbs and are not exported
        'PSUseApprovedVerbs'
        'PSUseSingularNouns'
        # Best-effort UI cleanup where a failure is expected and harmless; each has a comment
        'PSAvoidUsingEmptyCatchBlock'
        # Variables set inside event handler scriptblocks are read by other handlers
        'PSUseDeclaredVarsMoreThanAssignments'
        # Write-Log writes to the in-app log tab
        'PSAvoidOverwritingBuiltInCmdlets'
    )
}
