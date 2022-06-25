fn main() -> std::io::Result<()> {
    winres::WindowsResource::new()
        .set_manifest_file("manifest.xml")
        .compile()
}
