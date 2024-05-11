fn main() -> std::io::Result<()> {
    if cfg!(windows) {
        winres::WindowsResource::new()
            .set_manifest_file("manifest.xml")
            .compile()
    } else {
        Ok(())
    }
}
