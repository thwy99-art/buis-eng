import Foundation
import Capacitor
import Speech
import AVFoundation

@objc(SpeechPlugin)
public class SpeechPlugin: CAPPlugin, CAPBridgedPlugin {
    public let identifier = "SpeechPlugin"
    public let jsName = "SpeechPlugin"
    public let pluginMethods: [CAPPluginMethod] = [
        CAPPluginMethod(name: "requestPermission", returnType: CAPPluginReturnPromise),
        CAPPluginMethod(name: "startRecording", returnType: CAPPluginReturnPromise),
        CAPPluginMethod(name: "stopAndTranscribe", returnType: CAPPluginReturnPromise)
    ]

    private var audioEngine: AVAudioEngine = AVAudioEngine()
    private var recognitionRequest: SFSpeechAudioBufferRecognitionRequest?
    private var recognitionTask: SFSpeechRecognitionTask?
    private var recognizer: SFSpeechRecognizer?

    @objc func requestPermission(_ call: CAPPluginCall) {
        SFSpeechRecognizer.requestAuthorization { status in
            guard status == .authorized else {
                call.resolve(["granted": false])
                return
            }
            AVAudioSession.sharedInstance().requestRecordPermission { granted in
                call.resolve(["granted": granted])
            }
        }
    }

    @objc func startRecording(_ call: CAPPluginCall) {
        recognitionTask?.cancel()
        recognitionTask = nil
        if audioEngine.isRunning {
            audioEngine.stop()
        }
        audioEngine.inputNode.removeTap(onBus: 0)
        audioEngine = AVAudioEngine()

        let session = AVAudioSession.sharedInstance()
        do {
            try session.setCategory(.record, mode: .measurement, options: .duckOthers)
            try session.setActive(true, options: .notifyOthersOnDeactivation)
        } catch {
            call.reject("Audio session error: \(error.localizedDescription)")
            return
        }

        recognizer = SFSpeechRecognizer(locale: Locale(identifier: "en-US"))
        guard let recognizer = recognizer, recognizer.isAvailable else {
            call.reject("Speech recognizer unavailable")
            return
        }

        recognitionRequest = SFSpeechAudioBufferRecognitionRequest()
        guard let recognitionRequest = recognitionRequest else {
            call.reject("Cannot create recognition request")
            return
        }
        recognitionRequest.shouldReportPartialResults = false

        let inputNode = audioEngine.inputNode
        let format = inputNode.outputFormat(forBus: 0)
        inputNode.installTap(onBus: 0, bufferSize: 1024, format: format) { buffer, _ in
            recognitionRequest.append(buffer)
        }

        audioEngine.prepare()
        do {
            try audioEngine.start()
        } catch {
            call.reject("Audio engine error: \(error.localizedDescription)")
            return
        }

        call.resolve()
    }

    @objc func stopAndTranscribe(_ call: CAPPluginCall) {
        audioEngine.stop()
        audioEngine.inputNode.removeTap(onBus: 0)
        recognitionRequest?.endAudio()

        guard let recognizer = recognizer, let request = recognitionRequest else {
            call.reject("No active recording")
            return
        }

        recognitionTask = recognizer.recognitionTask(with: request) { [weak self] result, error in
            defer {
                self?.recognitionRequest = nil
                self?.recognitionTask = nil
                try? AVAudioSession.sharedInstance().setActive(false)
            }
            if let error = error {
                call.reject(error.localizedDescription)
                return
            }
            if let result = result, result.isFinal {
                call.resolve(["transcript": result.bestTranscription.formattedString])
            }
        }
    }
}
