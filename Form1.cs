// Created by: Justin Linwood Ross | 5/9/2025 |GitHub: https://github.com/Rythorian77?tab=repositories
// For Other Tutorials Go To: https://www.youtube.com/@justinlinwoodrossakarythor63/videos
// This application, 'Capture_Pro', is designed for real-time motion detection using a webcam.
// It captures video frames, applies image processing filters to detect motion, and can trigger alerts,
// save snapshots, and log events when motion is detected.
// The application is well-engineered to handle memory and resource management effectively.
// The consistent use of 'using' statements for disposable objects created within method scopes,
// coupled with explicit disposal of long-lived class members and proper shutdown procedures for external resources like the camera,
// significantly minimizes the risk of memory leaks.
// This provides a solid foundation for a stable application from a resource management perspective.

using AForge.Imaging;              // For image processing filters like Grayscale, Difference, Threshold, BlobCounter.
using AForge.Imaging.Filters;     // Specific namespace for AForge image filters.
using AForge.Video;               // Base namespace for AForge video capture.
using AForge.Video.DirectShow;    // For capturing video from DirectShow compatible devices (webcams).
using System;                     // Provides fundamental classes and base types.
using System.Collections.Generic; // For List<T>.
using System.Drawing;             // For graphics objects like Bitmap, Graphics, Pen, Rectangle, Color.
using System.Drawing.Imaging;     // For ImageFormat (used for saving images).
using System.IO;                  // For file and directory operations (Path, Directory, File).
using System.Media;               // For playing system sounds.
using System.Windows.Forms;       // For Windows Forms UI elements (Form, PictureBox, Button, ComboBox, Label, MessageBox).
using System.Diagnostics;         // Added for System.Diagnostics.EventLog and Debug.
using System.Speech.Synthesis;    // ADDED: For speech synthesis capabilities, alerting the user.

namespace Capture_Pro // This is the application's primary namespace. Consider renaming if 'Application\'s Name' is different.
{
    /// <summary>
    /// The main form for the Capture_Pro motion detection application.
    /// Handles camera initialization, frame processing, motion detection logic,
    /// and UI updates, ensuring robust resource management.
    /// </summary>
    public partial class Form1 : Form
    {
        #region Private Members - Variables

        #region Camera Related
        /// <summary>
        /// Collection of available video input devices (webcams) on the system.
        /// </summary>
        private FilterInfoCollection videoDevices;
        /// <summary>
        /// Represents the video capture device (webcam) currently in use.
        /// </summary>
        private VideoCaptureDevice videoSource;
        #endregion

        #region Motion Detection Related
        /// <summary>
        /// Used to count and extract information about detected blobs (connected components)
        /// in the motion map, representing areas of motion.
        /// </summary>
        private BlobCounter blobCounter;
        /// <summary>
        /// Stores the grayscale version of the previously processed frame.
        /// Used by the Difference filter to detect changes between consecutive frames.
        /// This object is disposed and re-assigned with each new frame.
        /// </summary>
        private Bitmap previousFrame;
        #endregion

        #region Filters
        /// <summary>
        /// A pre-initialized grayscale filter using the BT709 algorithm.
        /// This is a static instance as the filter itself is stateless and reusable.
        /// </summary>
        private readonly Grayscale grayscaleFilter = Grayscale.CommonAlgorithms.BT709;
        /// <summary>
        /// Filter used to calculate the absolute difference between the current frame
        /// and the previous frame, highlighting areas of change (motion).
        /// This object is re-initialized or its OverlayImage updated as previousFrame changes.
        /// </summary>
        private Difference differenceFilter;
        /// <summary>
        /// Filter used to convert the difference map into a binary image,
        /// where pixels above a certain threshold are white (motion) and others are black (no motion).
        /// The threshold value is configurable.
        /// </summary>
        private Threshold thresholdFilter;
        #endregion

        #region Drawing Pens
        /// <summary>
        /// Pen used for drawing rectangles around detected motion blobs on the video feed.
        /// Disposed explicitly on application shutdown or camera stop.
        /// </summary>
        private Pen greenPen = new Pen(Color.Green, 2);
        #endregion

        #region Automatic Enhancements & Settings
        /// <summary>
        /// Flag indicating whether active motion is currently being detected.
        /// </summary>
        private bool _isMotionActive = false;
        /// <summary>
        /// Timestamp of the last time an alert sound was played.
        /// Used to implement a cooldown period for alerts.
        /// </summary>
        private DateTime _lastAlertTime = DateTime.MinValue;
        /// <summary>
        /// Timestamp of the last time a motion snapshot was saved.
        /// Used to implement a cooldown period for saving images.
        /// </summary>
        private DateTime _lastSaveTime = DateTime.MinValue;
        /// <summary>
        /// Timestamp when the current continuous motion detection event started.
        /// Used for logging motion event durations.
        /// </summary>
        private DateTime _currentMotionStartTime = DateTime.MinValue;
        /// <summary>
        /// Timestamp of the very last frame where motion was detected.
        /// Used to determine when motion has ceased after a period of no detection.
        /// </summary>
        private DateTime _lastMotionDetectionTime = DateTime.MinValue;

        // Fixed Settings (No UI for adjustment in this version, can be made configurable later)
        /// <summary>
        /// The threshold value for the Threshold filter. Pixels with intensity difference
        /// above this value are considered motion. Lower values increase sensitivity.
        /// </summary>
        private readonly int _motionThreshold = 15;

        // Fixed Blob size filtering thresholds
        /// <summary>
        /// Minimum width (in pixels) a detected blob must have to be considered significant motion.
        /// </summary>
        private readonly int _minBlobWidth = 20;
        /// <summary>
        /// Minimum height (in pixels) a detected blob must have to be considered significant motion.
        /// </summary>
        private readonly int _minBlobHeight = 20;
        /// <summary>
        /// Maximum width (in pixels) a detected blob can have. Larger blobs might be noise or light changes.
        /// </summary>
        private readonly int _maxBlobWidth = 500;
        /// <summary>
        /// Maximum height (in pixels) a detected blob can have. Larger blobs might be noise or light changes.
        /// </summary>
        private readonly int _maxBlobHeight = 500;

        /// <summary>
        /// Cooldown duration after an alert sound is played before another can be played.
        /// Prevents continuous beeping during prolonged motion.
        /// </summary>
        private readonly TimeSpan _alertCooldown = TimeSpan.FromSeconds(5);
        /// <summary>
        /// Cooldown duration after an image is saved before another can be saved.
        /// Prevents an excessive number of images during prolonged motion.
        /// </summary>
        private readonly TimeSpan _saveCooldown = TimeSpan.FromSeconds(3);

        /// <summary>
        /// The directory path where detected motion images will be saved.
        /// Defaults to a 'Detected Images' folder on the user's Desktop.
        /// </summary>
        private readonly string _saveDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Detected Images");
        /// <summary>
        /// The file path for the application's motion event log.
        /// Located in the application's startup directory.
        /// </summary>
        private readonly string _logFilePath = Path.Combine(Application.StartupPath, "motion_log.txt");

        /// <summary>
        /// Flag to enable or disable saving motion event snapshots.
        /// </summary>
        private readonly bool _saveEventsEnabled = true;
        /// <summary>
        /// Flag to enable or disable playing an alert sound on motion detection.
        /// </summary>
        private readonly bool _alertSoundEnabled = true;
        /// <summary>
        /// Flag to enable or disable updating the status label on the UI for motion events.
        /// </summary>
        private readonly bool _alertLabelEnabled = true;
        /// <summary>
        /// Flag to enable or disable logging motion events to a file and EventLog.
        /// </summary>
        private readonly bool _logEventsEnabled = true;

        // --- Added for robust logging ---
        /// <summary>
        /// Counts consecutive failures when attempting to write to the motion log file.
        /// </summary>
        private int _consecutiveLogErrors = 0;
        /// <summary>
        /// The maximum number of consecutive log file write failures before considering
        /// file logging critically failed and potentially falling back to EventLog exclusively.
        /// </summary>
        private const int _maxConsecutiveLogErrors = 5;
        /// <summary>
        /// Flag to indicate if file logging has encountered a critical failure (e.g., permissions issues)
        /// beyond which it should not attempt to write to the file again.
        /// </summary>
        private bool _isFileLoggingCriticallyFailed = false;
        /// <summary>
        /// Timestamp of the last time a critical logging error message box was shown.
        /// Used to rate-limit message box pop-ups to avoid spamming the user.
        /// </summary>
        private DateTime _lastLogErrorMessageBoxTime = DateTime.MinValue;
        /// <summary>
        /// Cooldown duration for displaying a critical logging error message box.
        /// </summary>
        private readonly TimeSpan _logErrorMessageBoxCooldown = TimeSpan.FromMinutes(5);
        // --- End Added for robust logging ---

        #endregion

        #region Region of Interest (ROI) Variables
        /// <summary>
        /// Represents the selected Region of Interest (ROI) as a rectangle.
        /// If null, the entire frame is processed for motion.
        /// Initialized to cover the entire videoPictureBox by default.
        /// </summary>
        private Rectangle? _roiSelection = null;
        /// <summary>
        /// Pen used for drawing the Region of Interest rectangle on the video feed.
        /// Disposed explicitly on application shutdown or camera stop.
        /// </summary>
        private Pen _roiPen = new Pen(Color.Red, 2);
        #endregion

        #region Speech Synthesis // ADDED SECTION
        /// <summary>
        /// Speech synthesizer object for playing "Motion Alert" announcements.
        /// </summary>
        private SpeechSynthesizer _synthesizer;
        /// <summary>
        /// Timestamp of the last time a speech alert was played.
        /// </summary>
        private DateTime _lastSpeechAlertTime = DateTime.MinValue;
        /// <summary>
        /// Cooldown duration after a speech alert is played before another can be played.
        /// Set to 4 seconds as requested.
        /// </summary>
        private readonly TimeSpan _speechAlertCooldown = TimeSpan.FromSeconds(4);
        #endregion // END ADDED SECTION

        #endregion

        #region Constructor
        /// <summary>
        /// Initializes a new instance of the Form1 class.
        /// Sets up UI components, initializes motion detection objects,
        /// and loads available camera devices.
        /// </summary>
        public Form1()
        {
            InitializeComponent(); // Initializes components defined in the designer file (e.g., buttons, picture boxes).

            // Set the _roiSelection to match the initial size of the videoPictureBox.
            // This makes the entire video feed the default ROI for motion detection.
            _roiSelection = new Rectangle(0, 0, videoPictureBox.Width, videoPictureBox.Height);

            // Subscribe to the FormClosing event to ensure resources are properly released
            // when the application is closed by the user.
            FormClosing += new FormClosingEventHandler(Form1_FormClosing);

            // Initialize the BlobCounter, which is used to find and analyze motion blobs.
            blobCounter = new BlobCounter();
            // Initialize the Threshold filter with the default motion threshold.
            thresholdFilter = new Threshold(_motionThreshold);

            // Initially disable the Start and Stop buttons until cameras are loaded.
            startButton.Enabled = false;
            stopButton.Enabled = false;

            LoadCameraDevices(); // Attempt to discover and list available camera devices.

            // ADDED: Initialize speech synthesizer
            try
            {
                _synthesizer = new SpeechSynthesizer();
                bool voiceFound = false;
                foreach (InstalledVoice voice in _synthesizer.GetInstalledVoices())
                {
                    VoiceInfo info = voice.VoiceInfo;
                    // Look for voices with "en-GB" locale (UK English) and female gender
                    // The Culture.Name property usually follows the format "languagecode-countrycode" (e.g., "en-GB").
                    if (info.Culture.Name.Equals("en-GB", StringComparison.OrdinalIgnoreCase) && info.Gender == VoiceGender.Female)
                    {
                        _synthesizer.SelectVoice(info.Name);
                        voiceFound = true;
                        LogMotionEvent($"INFO: Selected speech voice: {info.Name} ({info.Culture.Name}, {info.Gender})");
                        break;
                    }
                }
                if (!voiceFound)
                {
                    // Fallback: Try to select any female voice if UK female is not found.
                    try
                    {
                        _synthesizer.SelectVoiceByHints(VoiceGender.Female);
                        LogMotionEvent("INFO: UK female voice not found, selected default female voice.");
                    }
                    catch
                    {
                        // Fallback: If no specific female voice found, use the default system voice.
                        LogMotionEvent("WARNING: No female voice found, using default available voice.");
                    }
                }
                _synthesizer.Volume = 100; // Set volume (0-100)
                _synthesizer.Rate = 0;     // Set speech rate (-10 to 10)
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error initializing speech synthesizer: {ex.Message}", "Speech Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                LogMotionEvent($"ERROR: Failed to initialize speech synthesizer: {ex.Message}");
                _synthesizer = null; // Ensure it's null if initialization fails
            }
            // END ADDED
        }
        #endregion

        #region Camera Management
        /// <summary>
        /// Discovers and populates the camera combo box with available video input devices.
        /// Handles cases where no cameras are found or an error occurs during discovery.
        /// </summary>
        private void LoadCameraDevices()
        {
            try
            {
                // Update status label on the UI, ensuring it's thread-safe via Invoke.
                if (statusLabel != null && statusLabel.InvokeRequired)
                {
                    statusLabel.Invoke((MethodInvoker)delegate { statusLabel.Text = "Loading cameras..."; });
                }
                else if (statusLabel != null)
                {
                    statusLabel.Text = "Loading cameras...";
                }

                // Create a collection of all video input devices found on the system.
                videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);

                if (videoDevices.Count == 0)
                {
                    // If no cameras are found, inform the user and disable the start button.
                    MessageBox.Show("No video sources found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    startButton.Enabled = false;
                    // Update status label.
                    if (statusLabel != null && statusLabel.InvokeRequired)
                    {
                        statusLabel.Invoke((MethodInvoker)delegate { statusLabel.Text = "No cameras found."; });
                    }
                    else if (statusLabel != null)
                    {
                        statusLabel.Text = "No cameras found.";
                    }
                    LogMotionEvent("WARNING: No video sources found."); // Log this significant event.
                }
                else
                {
                    // If cameras are found, clear any existing items and add each device's name to the combo box.
                    cameraComboBox.Items.Clear();
                    foreach (FilterInfo device in videoDevices)
                    {
                        cameraComboBox.Items.Add(device.Name);
                    }

                    // Select the first camera by default.
                    cameraComboBox.SelectedIndex = 0;
                    // Enable the Start button as a camera is available.
                    startButton.Enabled = true;

                    // Update status label with the number of cameras found.
                    if (statusLabel != null && statusLabel.InvokeRequired)
                    {
                        statusLabel.Invoke((MethodInvoker)delegate { statusLabel.Text = $"Found {videoDevices.Count} camera(s). Ready."; });
                    }
                    else if (statusLabel != null)
                    {
                        statusLabel.Text = $"Found {videoDevices.Count} camera(s). Ready.";
                    }
                    LogMotionEvent($"INFO: Found {videoDevices.Count} camera(s). Ready."); // Log successful camera detection.
                }
            }
            catch (Exception ex)
            {
                // Handle any exceptions during camera loading, display an error message,
                // and disable the start button to prevent further issues.
                MessageBox.Show("Failed to load video sources: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                startButton.Enabled = false;
                // Update status label.
                if (statusLabel != null && statusLabel.InvokeRequired)
                {
                    statusLabel.Invoke((MethodInvoker)delegate { statusLabel.Text = "Error loading cameras."; });
                }
                else if (statusLabel != null)
                {
                    statusLabel.Text = "Error loading cameras.";
                }
                LogMotionEvent($"ERROR: Failed to load video sources: {ex.Message}"); // Log the error for debugging.
            }
        }

        /// <summary>
        /// Stops the currently running video source and disposes of all associated resources
        /// to prevent memory leaks and ensure a clean shutdown. This method is designed
        /// to be robust against individual disposal failures.
        /// </summary>
        private void StopCamera()
        {
            // Attempt to stop the video source first.
            try
            {
                if (videoSource != null && videoSource.IsRunning)
                {
                    // Unsubscribe from the NewFrame event to prevent further processing
                    // after the camera has stopped.
                    videoSource.NewFrame -= new NewFrameEventHandler(VideoSource_NewFrame);
                    // Signal the video source to stop capturing frames.
                    videoSource.SignalToStop();
                    // Wait for the video source to completely stop. This is crucial for proper shutdown.
                    videoSource.WaitForStop();
                    // Nullify the videoSource reference to allow garbage collection.
                    videoSource = null;
                    LogMotionEvent("INFO: Video source signaled to stop and waited for stop.");
                }
            }
            catch (Exception ex)
            {
                LogMotionEvent($"ERROR: Exception while stopping video source: {ex.Message}");
                // Attempt to continue with resource disposal even if stopping the video source
                // failed partially, as other resources might still need to be released.
            }

            // Centralized disposal logic with individual try-catch blocks for robustness.
            // This ensures that if one resource fails to dispose, others can still be released.
            try
            {
                if (videoPictureBox.Image != null)
                {
                    // Dispose the image currently displayed in the PictureBox to release its memory.
                    videoPictureBox.Image.Dispose();
                    videoPictureBox.Image = null; // Clear the image reference.
                }
            }
            catch (Exception ex)
            {
                LogMotionEvent($"ERROR: Failed to dispose videoPictureBox.Image: {ex.Message}");
            }

            try
            {
                if (previousFrame != null)
                {
                    // Dispose the previous frame bitmap used for motion detection.
                    previousFrame.Dispose();
                    previousFrame = null; // Clear the reference.
                }
            }
            catch (Exception ex)
            {
                LogMotionEvent($"ERROR: Failed to dispose previousFrame: {ex.Message}");
            }

            // AForge filter objects like Difference and Threshold typically don't hold
            // unmanaged resources that require explicit Dispose(). They are usually
            // re-initialized or their properties updated. Nullifying them here
            // helps with garbage collection and ensures a clean state reset for restart.
            try
            {
                if (differenceFilter != null)
                {
                    differenceFilter = null;
                }
            }
            catch (Exception ex) // Catching just in case, though unlikely for nulling references
            {
                LogMotionEvent($"ERROR: Failed to reset differenceFilter: {ex.Message}");
            }

            try
            {
                if (thresholdFilter != null)
                {
                    thresholdFilter = null;
                }
            }
            catch (Exception ex)
            {
                LogMotionEvent($"ERROR: Failed to reset thresholdFilter: {ex.Message}");
            }

            try
            {
                if (blobCounter != null)
                {
                    // BlobCounter also typically doesn't require explicit dispose,
                    // but nulling helps for state reset.
                    blobCounter = null;
                }
            }
            catch (Exception ex)
            {
                LogMotionEvent($"ERROR: Failed to reset blobCounter: {ex.Message}");
            }

            // Dispose of GDI+ objects (Pens). These *do* hold unmanaged resources and must be disposed.
            try
            {
                if (greenPen != null)
                {
                    greenPen.Dispose();
                    greenPen = null; // Clear the reference.
                }
                // Re-initialize the pen for subsequent starts if needed, or ensure it's created in the constructor.
                // For consistency, if it's a class member, it's better to recreate it than to leave it null.
                greenPen = new Pen(Color.Green, 2); // Re-create for next use
            }
            catch (Exception ex)
            {
                LogMotionEvent($"ERROR: Failed to dispose or re-initialize greenPen: {ex.Message}");
            }

            try
            {
                if (_roiPen != null)
                {
                    _roiPen.Dispose();
                    _roiPen = null; // Clear the reference.
                }
                _roiPen = new Pen(Color.Red, 2); // Re-create for next use
            }
            catch (Exception ex)
            {
                 LogMotionEvent($"ERROR: Failed to dispose or re-initialize _roiPen: {ex.Message}");
            }

            // Reset motion detection state flags.
            _isMotionActive = false;

            // Update UI status label.
            if (statusLabel != null && statusLabel.InvokeRequired)
            {
                statusLabel.Invoke((MethodInvoker)delegate { statusLabel.Text = "Camera stopped. Ready."; });
            }
            else if (statusLabel != null)
            {
                statusLabel.Text = "Camera stopped. Ready.";
            }

            // Update button and combo box states for restarting.
            startButton.Enabled = true;
            stopButton.Enabled = false;
            cameraComboBox.Enabled = true;

            LogMotionEvent("INFO: Camera stopped. Resources disposed and reset.");
        }
        #endregion

        #region Event Handlers
        /// <summary>
        /// Handles the click event for the 'Start' button.
        /// Initializes and starts the selected video capture device.
        /// Resets motion detection state variables.
        /// </summary>
        /// <param name="sender">The object that raised the event.</param>
        /// <param name="e">The event data.</param>
        private void StartButton_Click(object sender, EventArgs e)
        {
            try
            {
                // Validate that a camera is selected and available.
                if (videoDevices == null || videoDevices.Count == 0 || cameraComboBox.SelectedIndex < 0)
                {
                    MessageBox.Show("Please select a camera.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    if (statusLabel != null) statusLabel.Text = "Select a camera.";
                    LogMotionEvent("WARNING: Start button clicked with no camera selected or found.");
                    return; // Exit the method if no camera is selected.
                }

                // Re-initialize threshold filter if it's null or its value has changed
                // (though _motionThreshold is a readonly field, this check is good practice if it were configurable).
                if (thresholdFilter == null || thresholdFilter.ThresholdValue != _motionThreshold)
                {
                    thresholdFilter = new Threshold(_motionThreshold);
                }

                // Create a new VideoCaptureDevice instance using the moniker string of the selected device.
                videoSource = new VideoCaptureDevice(videoDevices[cameraComboBox.SelectedIndex].MonikerString);
                // Unsubscribe defensively to prevent multiple subscriptions if Start is clicked repeatedly
                // without a full application restart or proper StopCamera call.
                videoSource.NewFrame -= new NewFrameEventHandler(VideoSource_NewFrame);
                // Subscribe to the NewFrame event, which is triggered whenever a new video frame is available.
                videoSource.NewFrame += new NewFrameEventHandler(VideoSource_NewFrame);
                // Begin video capture.
                videoSource.Start();

                // Dispose of any previous frame stored from a prior run and clear the reference.
                if (previousFrame != null)
                {
                    previousFrame.Dispose();
                    previousFrame = null;
                }
                // Reset the difference filter as the previous frame has been reset.
                differenceFilter = null;

                // Reset all motion detection state variables to their initial values.
                _lastAlertTime = DateTime.MinValue;
                _lastSaveTime = DateTime.MinValue;
                _isMotionActive = false;
                _currentMotionStartTime = DateTime.MinValue;
                _lastMotionDetectionTime = DateTime.MinValue;
                _lastSpeechAlertTime = DateTime.MinValue; // ADDED: Reset speech alert time on start

                // Update UI element states.
                startButton.Enabled = false;
                stopButton.Enabled = true;
                cameraComboBox.Enabled = false;

                // Update status label.
                if (_alertLabelEnabled && statusLabel != null) statusLabel.Text = "Camera started. Detecting motion...";
                LogMotionEvent("INFO: Application started. Camera activated.");
            }
            catch (Exception ex)
            {
                // Handle any errors during camera startup.
                MessageBox.Show("Error starting camera: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // Re-enable start button and disable stop button on error.
                startButton.Enabled = true;
                stopButton.Enabled = false;
                cameraComboBox.Enabled = true;
                // Attempt to signal the video source to stop if it somehow started partially.
                if (videoSource != null && videoSource.IsRunning)
                {
                    videoSource.SignalToStop();
                }
                videoSource = null; // Ensure videoSource reference is cleared.
                if (_alertLabelEnabled && statusLabel != null) statusLabel.Text = "Error starting camera.";
                LogMotionEvent($"ERROR: Failed to start camera: {ex.Message}"); // Log the error.
            }
        }

        /// <summary>
        /// Event handler for when a new frame is received from the video source.
        /// This method is crucial for real-time processing and motion detection.
        /// It clones the frame, processes it for motion, updates the UI, and handles
        /// motion-triggered events (alerts, saving snapshots).
        /// </summary>
        /// <param name="sender">The video source that sent the frame.</param>
        /// <param name="eventArgs">Arguments containing the new frame.</param>
        private void VideoSource_NewFrame(object sender, NewFrameEventArgs eventArgs)
        {
            // 'currentFrame' will hold a clone of the original frame for display.
            Bitmap currentFrame = null;
            // 'frameToProcess' will hold a clone specifically for motion analysis,
            // allowing drawing on 'currentFrame' for display without affecting analysis.
            Bitmap frameToProcess = null;

            try
            {
                // Clone the incoming frame to work with it.
                // The original eventArgs.Frame is managed by AForge and should not be disposed here.
                currentFrame = (Bitmap)eventArgs.Frame.Clone();
                // Clone again for the motion processing pipeline.
                frameToProcess = (Bitmap)currentFrame.Clone();

                // Process the 'frameToProcess' for motion detection.
                // This method will draw rectangles on 'frameToProcess' if motion is found.
                bool motionDetected = ProcessFrameForMotion(frameToProcess);

                // Update the UI (videoPictureBox and statusLabel) on the UI thread.
                // This is critical because NewFrame events occur on a separate thread.
                if (videoPictureBox.InvokeRequired)
                {
                    videoPictureBox.Invoke((MethodInvoker)delegate
                    {
                        try
                        {
                            // Dispose of the previously displayed image in the PictureBox
                            // to prevent GDI+ memory leaks.
                            videoPictureBox.Image?.Dispose();
                            // Assign the processed frame (which might have motion rectangles) to the PictureBox.
                            videoPictureBox.Image = frameToProcess;
                            // Handle motion-related events (alerts, saving) using the original frame for saving.
                            HandleMotionEvents(motionDetected, currentFrame); // Pass the original clone for saving.
                        }
                        catch (Exception invokeEx)
                        {
                            // Log errors occurring within the Invoke delegate.
                            System.Diagnostics.Debug.WriteLine($"Error in VideoPictureBox Invoke delegate: {invokeEx.Message}");
                            LogMotionEvent($"ERROR: Exception in VideoPictureBox Invoke delegate: {invokeEx.Message}");
                            // Ensure disposal of bitmaps if an error occurs during UI update.
                            currentFrame?.Dispose();
                            frameToProcess?.Dispose();
                            // Attempt to stop the camera on a critical error that prevents UI updates.
                            StopCamera();
                        }
                    });
                }
                else // If not invoked (should generally not happen for NewFrame, but as fallback).
                {
                    videoPictureBox.Image?.Dispose();
                    videoPictureBox.Image = frameToProcess;
                    HandleMotionEvents(motionDetected, currentFrame);
                }
            }
            catch (Exception ex)
            {
                // Log errors occurring in the main NewFrame handler logic.
                System.Diagnostics.Debug.WriteLine($"Error in NewFrame (main thread logic): {ex.Message}");
                LogMotionEvent($"ERROR: Exception in NewFrame handler: {ex.Message}");
                // Ensure all created bitmaps are disposed in case of an error.
                currentFrame?.Dispose();
                currentFrame = null;
                frameToProcess?.Dispose();
                frameToProcess = null;
                // Attempt to stop the camera on a critical error that prevents further frames from being processed.
                StopCamera();
            }
        }

        /// <summary>
        /// Handles the click event for the 'Stop' button.
        /// Initiates the process of stopping the camera and releasing associated resources.
        /// </summary>
        /// <param name="sender">The object that raised the event.</param>
        /// <param name="e">The event data.</param>
        private void StopButton_Click(object sender, EventArgs e)
        {
            StopCamera(); // Call the centralized StopCamera method.
        }

        /// <summary>
        /// Event handler for the FormClosing event.
        /// Ensures that the camera is stopped and all resources are released
        /// gracefully when the application is closed.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">A FormClosingEventArgs that contains the event data.</param>
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            StopCamera(); // Ensure camera is stopped and resources released.

            // ADDED: Dispose of speech synthesizer
            try
            {
                if (_synthesizer != null)
                {
                    _synthesizer.Dispose();
                    _synthesizer = null;
                    LogMotionEvent("INFO: Speech synthesizer disposed.");
                }
            }
            catch (Exception ex)
            {
                LogMotionEvent($"ERROR: Failed to dispose speech synthesizer: {ex.Message}");
            }
            // END ADDED
        }

        #endregion

        #region Motion Detection Logic
        /// <summary>
        /// Processes a single video frame to detect motion.
        /// Applies grayscale, difference, and threshold filters, then uses BlobCounter
        /// to identify significant motion areas. Draws rectangles around detected motion.
        /// </summary>
        /// <param name="image">The bitmap image representing the current frame to process.</param>
        /// <returns>True if significant motion is detected, false otherwise.</returns>
        private bool ProcessFrameForMotion(Bitmap image)
        {
            bool motionDetected = false;

            try
            {
                // Ensure image is in 24bppRgb format for AForge.NET filters if it's not already.
                if (image.PixelFormat != PixelFormat.Format24bppRgb)
                {
                    using (Bitmap tempImage = AForge.Imaging.Image.Clone(image, PixelFormat.Format24bppRgb))
                    {
                        image.Dispose(); // Dispose the original image since we're replacing it.
                        image = tempImage; // Assign the 24bppRgb clone.
                    }
                }

                // Apply Grayscale filter to convert the frame to grayscale.
                // This is necessary for the Difference filter to work correctly.
                using (Bitmap grayFrame = grayscaleFilter.Apply(image))
                {
                    // Initialize the difference filter with the first frame.
                    if (previousFrame == null)
                    {
                        previousFrame = (Bitmap)grayFrame.Clone();
                        // No motion can be detected on the first frame.
                        return false;
                    }

                    // Initialize or update the Difference filter.
                    // We create a new instance if it's null or if the previousFrame has changed size.
                    // (Though in this application, previousFrame is consistently the same size as grayFrame)
                    if (differenceFilter == null)
                    {
                        differenceFilter = new Difference(previousFrame);
                    }
                    else
                    {
                        // Ensure the Difference filter's overlay image is the current previous frame.
                        // This is crucial if previousFrame was disposed and re-assigned.
                        differenceFilter.OverlayImage = previousFrame;
                    }

                    // Apply the Difference filter to get the difference between the current and previous grayscale frames.
                    using (Bitmap diffFrame = differenceFilter.Apply(grayFrame))
                    {
                        // Apply the Threshold filter to convert the difference image into a binary image.
                        // Pixels with a difference above _motionThreshold will be white (motion), others black.
                        thresholdFilter.ApplyInPlace(diffFrame);

                        // The BlobCounter needs to process the binary image (diffFrame)
                        // to identify the blobs before GetObjectsInformation() can retrieve them.
                        blobCounter.ProcessImage(diffFrame);

                        // Get information about objects (blobs) in the binary difference image.
                        // This method correctly takes no arguments and returns an array of Blob objects.
                        Blob[] blobs = blobCounter.GetObjectsInformation(); // <-- CORRECTED LINE

                        // Filter out small blobs and draw rectangles around significant motion.
                        List<Rectangle> motionRectangles = new List<Rectangle>();
                        foreach (Blob blob in blobs)
                        {
                            // Filter blobs by size to ignore noise or minor changes.
                            if (blob.Rectangle.Width >= _minBlobWidth &&
                                blob.Rectangle.Height >= _minBlobHeight &&
                                blob.Rectangle.Width <= _maxBlobWidth &&
                                blob.Rectangle.Height <= _maxBlobHeight)
                            {
                                // Check if the blob is within the defined Region of Interest (ROI).
                                if (_roiSelection.HasValue && !_roiSelection.Value.IntersectsWith(blob.Rectangle))
                                {
                                    continue; // Skip blobs outside the ROI.
                                }

                                motionDetected = true;
                                motionRectangles.Add(blob.Rectangle);
                            }
                        }

                        // Draw rectangles around detected motion on the original 'image' (currentFrame clone).
                        if (motionDetected)
                        {
                            using (Graphics g = Graphics.FromImage(image))
                            {
                                foreach (Rectangle rect in motionRectangles)
                                {
                                    g.DrawRectangle(greenPen, rect);
                                }
                            }
                        }

                        // Dispose the previous frame and set the current grayscale frame as the new previous frame.
                        previousFrame?.Dispose();
                        previousFrame = (Bitmap)grayFrame.Clone();
                    }
                }
            }
            catch (OutOfMemoryException oomEx)
            {
                // Handle out-of-memory errors which can occur if frames are too large
                // or processing is too slow causing a queue buildup.
                LogMotionEvent($"CRITICAL ERROR: Out of memory during frame processing: {oomEx.Message}. Stopping camera.");
                MessageBox.Show("System ran out of memory during image processing. Stopping camera.", "Critical Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                StopCamera();
                return false;
            }
            catch (Exception ex)
            {
                // Log any other exceptions during frame processing.
                LogMotionEvent($"ERROR: Exception during frame processing: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Frame Processing Error: {ex.Message}");
                // Consider stopping the camera if errors are persistent.
            }

            return motionDetected;
        }
        #endregion

        #region Motion Event Handling & Logging
        /// <summary>
        /// Handles actions to be taken when motion is detected or motion ceases.
        /// This includes playing alerts, saving snapshots, and updating UI status.
        /// </summary>
        /// <param name="motionDetected">True if motion is currently detected, false otherwise.</param>
        /// <param name="currentFrame">The current video frame (original, before motion drawing) to be saved if a snapshot is triggered.</param>
        private void HandleMotionEvents(bool motionDetected, Bitmap currentFrame)
        {
            // Ensure the status label is updated on the UI thread.
            // This method is typically called from a BeginInvoke/Invoke block,
            // so direct access to statusLabel.Text is usually safe here,
            // but a check is added for robustness.
            if (statusLabel != null)
            {
                if (motionDetected)
                {
                    _lastMotionDetectionTime = DateTime.Now; // Update the last time motion was truly detected.

                    // If motion just started (was previously inactive)
                    if (!_isMotionActive)
                    {
                        _isMotionActive = true;
                        _currentMotionStartTime = DateTime.Now;
                        if (_alertLabelEnabled) statusLabel.Text = "MOTION DETECTED!";
                        LogMotionEvent("ALERT: Motion event started.");
                    }
                    else
                    {
                        // Motion is ongoing
                        if (_alertLabelEnabled) statusLabel.Text = "MOTION ACTIVE...";
                    }

                    // Play alert sound if enabled and cooldown has passed.
                    if (_alertSoundEnabled && (DateTime.Now - _lastAlertTime > _alertCooldown))
                    {
                        SystemSounds.Beep.Play();
                        _lastAlertTime = DateTime.Now;
                        LogMotionEvent("INFO: Alert sound played.");
                    }

                    // ADDED: Play speech alert if enabled and cooldown has passed.
                    if (_synthesizer != null && (DateTime.Now - _lastSpeechAlertTime > _speechAlertCooldown))
                    {
                        try
                        {
                            _synthesizer.SpeakAsync("Motion Alert");
                            _lastSpeechAlertTime = DateTime.Now;
                            LogMotionEvent("INFO: Speech alert 'Motion Alert' played.");
                        }
                        catch (Exception ex)
                        {
                            LogMotionEvent($"ERROR: Failed to play speech alert: {ex.Message}");
                        }
                    }
                    // END ADDED

                    // Save snapshot if enabled and cooldown has passed.
                    if (_saveEventsEnabled && (DateTime.Now - _lastSaveTime > _saveCooldown))
                    {
                        SaveSnapshot(currentFrame);
                        _lastSaveTime = DateTime.Now;
                        LogMotionEvent("INFO: Motion snapshot saved.");
                    }
                }
                else // No motion detected in the current frame
                {
                    // If motion was active and now there's a lull (no detection for a short period)
                    if (_isMotionActive && (DateTime.Now - _lastMotionDetectionTime > TimeSpan.FromSeconds(1))) // 1-second grace period for ending motion
                    {
                        TimeSpan motionDuration = DateTime.Now - _currentMotionStartTime;
                        if (_alertLabelEnabled) statusLabel.Text = "No motion. Camera active.";
                        LogMotionEvent($"INFO: Motion event ended. Duration: {motionDuration.TotalSeconds:F2} seconds.");
                        _isMotionActive = false;
                    }
                    else if (!_isMotionActive)
                    {
                        if (_alertLabelEnabled) statusLabel.Text = "No motion. Camera active.";
                    }
                }
            }
        }

        /// <summary>
        /// Saves a snapshot of the current video frame to a designated directory.
        /// The filename includes a timestamp for uniqueness.
        /// Robustly handles directory creation and file writing errors.
        /// </summary>
        /// <param name="image">The bitmap image to be saved.</param>
        private void SaveSnapshot(Bitmap image)
        {
            if (!_saveEventsEnabled) return; // Do nothing if saving is disabled.

            try
            {
                // Ensure the save directory exists, create it if it doesn't.
                if (!Directory.Exists(_saveDirectory))
                {
                    Directory.CreateDirectory(_saveDirectory);
                    LogMotionEvent($"INFO: Created save directory: {_saveDirectory}");
                }

                // Generate a unique filename with a timestamp.
                string filename = $"Motion_{DateTime.Now:yyyyMMdd_HHmmssfff}.jpg";
                string filePath = Path.Combine(_saveDirectory, filename);

                // Save the image in JPEG format.
                // Use a 'using' statement to ensure the Bitmap object is properly disposed.
                // Cloning the image before saving is crucial because the original 'currentFrame'
                // might still be in use by the display or other processes, and disposing it here
                // would cause issues.
                using (Bitmap imageToSave = (Bitmap)image.Clone())
                {
                    imageToSave.Save(filePath, ImageFormat.Jpeg);
                }
                LogMotionEvent($"INFO: Snapshot saved to: {filePath}");
            }
            catch (UnauthorizedAccessException uaEx)
            {
                LogMotionEvent($"ERROR: Access denied when saving snapshot to {_saveDirectory}: {uaEx.Message}");
                // Potentially inform user, but avoid message box spam.
                if (DateTime.Now - _lastLogErrorMessageBoxTime > _logErrorMessageBoxCooldown)
                {
                    MessageBox.Show($"Cannot save images to '{_saveDirectory}'. Please check permissions. Error: {uaEx.Message}",
                        "Permission Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    _lastLogErrorMessageBoxTime = DateTime.Now;
                }
            }
            catch (Exception ex)
            {
                // Catch any other exceptions during file saving.
                LogMotionEvent($"ERROR: Failed to save snapshot: {ex.Message}");
                if (DateTime.Now - _lastLogErrorMessageBoxTime > _logErrorMessageBoxCooldown)
                {
                    MessageBox.Show($"An error occurred while saving a snapshot: {ex.Message}",
                        "Save Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    _lastLogErrorMessageBoxTime = DateTime.Now;
                }
            }
        }

        /// <summary>
        /// Logs important motion detection events and errors to a file and the Windows Event Log.
        /// Implements robust error handling for file operations and rate-limits message boxes.
        /// </summary>
        /// <param name="message">The message to log.</param>
        private void LogMotionEvent(string message)
        {
            if (!_logEventsEnabled) return; // Do nothing if logging is disabled.

            string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} - {message}";

            // Attempt to log to file first, if not critically failed.
            if (!_isFileLoggingCriticallyFailed)
            {
                try
                {
                    // Ensure the directory exists.
                    string logDirectory = Path.GetDirectoryName(_logFilePath);
                    if (!string.IsNullOrEmpty(logDirectory) && !Directory.Exists(logDirectory))
                    {
                        Directory.CreateDirectory(logDirectory);
                    }

                    File.AppendAllText(_logFilePath, logEntry + Environment.NewLine);
                    _consecutiveLogErrors = 0; // Reset error counter on success.
                }
                catch (Exception fileEx)
                {
                    _consecutiveLogErrors++;
                    System.Diagnostics.Debug.WriteLine($"File logging error (attempt {_consecutiveLogErrors}): {fileEx.Message}");

                    // If consecutive failures exceed threshold, mark file logging as critically failed.
                    if (_consecutiveLogErrors >= _maxConsecutiveLogErrors)
                    {
                        _isFileLoggingCriticallyFailed = true;
                        LogToEventLog($"CRITICAL FILE LOGGING ERROR: File logging disabled due to persistent errors: {fileEx.Message}", EventLogEntryType.Error);
                        if (DateTime.Now - _lastLogErrorMessageBoxTime > _logErrorMessageBoxCooldown)
                        {
                            MessageBox.Show($"Persistent errors writing to log file '{_logFilePath}'. File logging will be disabled. Error: {fileEx.Message}",
                                "Logging Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            _lastLogErrorMessageBoxTime = DateTime.Now;
                        }
                    }
                    else
                    {
                        // Log to EventLog for temporary file issues, but don't disable file logging yet.
                        LogToEventLog($"WARNING: Temporary file logging error to {_logFilePath}: {fileEx.Message}", EventLogEntryType.Warning);
                    }
                }
            }

            // Always attempt to log to EventLog as a fallback and for critical application events.
            // This also serves as the primary log destination if file logging is critically failed.
            if (_isFileLoggingCriticallyFailed || message.StartsWith("ERROR") || message.StartsWith("CRITICAL ERROR") || message.StartsWith("WARNING") || message.StartsWith("ALERT"))
            {
                EventLogEntryType entryType = EventLogEntryType.Information;
                if (message.StartsWith("ERROR") || message.StartsWith("CRITICAL ERROR"))
                {
                    entryType = EventLogEntryType.Error;
                }
                else if (message.StartsWith("WARNING") || message.StartsWith("ALERT"))
                {
                    entryType = EventLogEntryType.Warning;
                }
                LogToEventLog(logEntry, entryType);
            }

            // Also output to debug console (for development).
            System.Diagnostics.Debug.WriteLine(logEntry);
        }

        /// <summary>
        /// Helper method to log messages to the Windows Event Log.
        /// </summary>
        /// <param name="message">The message to log.</param>
        /// <param name="type">The type of event log entry.</param>
        private void LogToEventLog(string message, EventLogEntryType type)
        {
            try
            {
                string source = "Capture_Pro";
                if (!EventLog.SourceExists(source))
                {
                    EventLog.CreateEventSource(source, "Application");
                }
                EventLog.WriteEntry(source, message, type);
            }
            catch (Exception ex)
            {
                // Fallback if EventLog itself fails (e.g., permissions issue for creating source).
                System.Diagnostics.Debug.WriteLine($"CRITICAL: Failed to write to EventLog: {ex.Message} - Message: {message}");
                if (DateTime.Now - _lastLogErrorMessageBoxTime > _logErrorMessageBoxCooldown)
                {
                    MessageBox.Show($"Critical error: Failed to write to Windows Event Log. Error: {ex.Message}",
                        "Event Log Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    _lastLogErrorMessageBoxTime = DateTime.Now;
                }
            }
        }

        #endregion

        #region Region of Interest (ROI) Drawing Logic
        /// <summary>
        /// Handles the Paint event for the PictureBox to draw the ROI rectangle.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">A PaintEventArgs that contains the event data.</param>
        private void VideoPictureBox_Paint(object sender, PaintEventArgs e)
        {
            // Draw the ROI rectangle if it has been defined.
            if (_roiSelection.HasValue)
            {
                e.Graphics.DrawRectangle(_roiPen, _roiSelection.Value);
            }
        }
        #endregion
    }
}