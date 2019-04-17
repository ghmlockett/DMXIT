# DMXIT - Prototype project only.  Work in progress

CODE OWNED BY MIKE LOCKETT AND AVAILABLE FOR DEVELOPERS TO REFERENCE FOR LEARNING PURPSOSES
NOT FOR PRODUCTION DISTRIBUTION WITHOUT OWNER CONSENT
lockettm@msn.com | 952.292.1827


The project requires MongoDB to be installed.  Follow these steps to install Mongo DB:

1. Download 4.0.9 from https://www.mongodb.com/download-center/community
2. Install the MSI package
3. Recommend also installing Robo Mongo 3T from https://robomongo.org/ to view and administer MongoDB DB and collections

The project requires the following DMX USB interface:
https://dmxking.com/usbdmx/ultradmxmicro

Once compiled
1. Use the FIXTURES tab to store DMX devices
2. Use the FIXTURES tab to configure the DMX devices with their respective channel functionality
    Note: Configure each channel and channel functionality.  
    Note: List same channel multiple times with respective starting value if more than one function is available in a single channel
3. Use the LAYOUTS tab to create new layout and assign fixtures to the layout (set the respective DMX starting channels here)
4. Use the BOARD tab to control the DMX fixtures in real-time and save/run slices

Note: The following tabs are just for testing and not ready (Manual control, test, scenes)
Note: The Video tab loads just to files at the moment (hard-coded) to demonstrate how to call slices at specific locations in the media
Note: While playing the media, click "MARK" where a slice should be called.  Then edit the markings to call the intended slice.
Note: The slices will be played back at the specified mark time when the media plays again.


