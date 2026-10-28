import AppKit
import CoreServices
import Foundation
import UniformTypeIdentifiers

let bundle = URL(fileURLWithPath: CommandLine.arguments[1]).standardizedFileURL
let files = CommandLine.arguments.dropFirst(2).map { URL(fileURLWithPath: $0) }
let workspace = NSWorkspace.shared
// A fresh CI Mac can have only an implicit TextEdit fallback for CSV. A newly
// registered editor may legitimately replace that fallback without changing
// any user preference. Seed explicit choices in this disposable CI account to
// test that actual user selections survive candidate registration.
let types = files.map { UTType(filenameExtension: $0.pathExtension)!.identifier }
let selectedApp = "com.apple.TextEdit"
for type in types {
    guard LSSetDefaultRoleHandlerForContentType(type as CFString, .all, selectedApp as CFString) == noErr else {
        fatalError("Could not seed explicit CI default for \(type)")
    }
}
guard LSRegisterURL(bundle as CFURL, true) == noErr else { fatalError("Launch Services registration failed") }
var result = [[String: Any]]()
for (index, file) in files.enumerated() {
    let candidates = workspace.urlsForApplications(toOpen: file).map { $0.standardizedFileURL }
    guard candidates.contains(bundle) else { fatalError("Missing Open With candidate for \(file.pathExtension): \(candidates)") }
    let selected = LSCopyDefaultRoleHandlerForContentType(types[index] as CFString, .all)?.takeRetainedValue() as String?
    guard selected?.lowercased() == selectedApp.lowercased() else {
        fatalError("Changed explicit default for \(file.pathExtension): \(selected ?? "none")")
    }
    result.append(["extension": file.pathExtension, "offered": true, "defaultPreserved": true])
}
let data = try JSONSerialization.data(withJSONObject: result, options: [.prettyPrinted, .sortedKeys])
print(String(data: data, encoding: .utf8)!)
