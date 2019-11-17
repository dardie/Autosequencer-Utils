# AutoSequencer-Utils

A powershell module to make working with the Microsoft AppV Autosequencer more Joyful.

The useful functions are:
* **New-SequencerPSSession** : Create PSSession to Autosequencer VM
* **Enter-SequencerPSSession** : Enter a PSSession on the Sequencer VM
* **Connect-Sequencer** : Create a terminal session to Autosequencer VM (parameters H, W, Explorer, PS)
* **Get-SequencerLogin** : Create a terminal session to Autosequencer VM
* **Add-SequencerToHosts** : Add Sequencer VM to hosts and fix any network issues that might affect connecting to it
* **Reset-Sequencer** : Reset sequencer to checkpoint *sequencer-base*
* **Test-AppV** : Automatically find a .appV file in current folder or subfolder, and publish/unpublish it, for testing purposes

If there is only a single autosequencer VM on localhost, the commands above can be invoked with no arguments.
If there is more than one autosequencer VM, they take an optional VMName parameter to specify which autosequencer to invoke.
    
