# AutoSequencer-Utils

A powershell module to make working with the Microsoft AppV Autosequencer more Joyful.

The useful functions are:
* **New-SequencerPSSession** : Create a PSSession to Autosequencer VM
* **Enter-SequencerPSSession** : Enter an interactive PSSession on the Sequencer VM
* **Connect-Sequencer** : Create a terminal session to Autosequencer VM (parameters H, W, Explorer, PS)
* **Get-SequencerLogin** : Retrieve the generated login credentials for the Autosequencer VM
* **Add-SequencerToHosts** : Add Sequencer VM to hosts file and invoke some tweaks to enable and optimize connecting to the VM
* **Reset-Sequencer** : Reset sequencer to checkpoint *sequencer-base*
* **Test-AppV** : Automatically find a .appV file in the current folder or subfolder and publish/unpublish it, for testing purposes

If there is only a single autosequencer VM on localhost, the commands above can be invoked with no arguments.
If there is more than one autosequencer VM, they take an optional VMName parameter to specify which autosequencer to invoke.
