import AppKit
import CoreServices
import Foundation

let bundle = URL(fileURLWithPath: CommandLine.arguments[1]).standardizedFileURL
let files = CommandLine.arguments.dropFirst(2).map { URL(fileURLWithPath: $0) }
let workspace = NSWorkspace.shared
let defaults = files.map { workspace.urlForApplication(toOpen: $0)?.standardizedFileURL }
guard LSRegisterURL(bundle as CFURL, true) == noErr else { fatalError("Launch Services registration failed") }
var result = [[String: Any]]()
for (index, file) in files.enumerated() {
    let candidates = workspace.urlsForApplications(toOpen: file).map { $0.standardizedFileURL }
    guard candidates.contains(bundle) else { fatalError("Missing Open With candidate for \(file.pathExtension): \(candidates)") }
    if let previous = defaults[index] {
        guard workspace.urlForApplication(toOpen: file)?.standardizedFileURL == previous else {
            fatalError("Changed existing default for \(file.pathExtension)")
        }
    }
    result.append(["extension": file.pathExtension, "offered": true, "defaultPreserved": true])
}
let data = try JSONSerialization.data(withJSONObject: result, options: [.prettyPrinted, .sortedKeys])
print(String(data: data, encoding: .utf8)!)
