#![windows_subsystem = "windows"]
use druid::{AppLauncher, WindowDesc, Widget, widget::{Label, Flex, Button}, LocalizedString, Data, Lens};
use pyo3::prelude::*;
use pyo3::types::IntoPyDict;

const VERTICAL_WIDGET_SPACING: f64 = 20.0;
const TEXT_BOX_WIDTH: f64 = 200.0;
const WINDOWS_TITLE: LocalizedString<HelloState> = LocalizedString::new("VAL");

#[derive(Clone, Data, Lens)]
struct HelloState {
    name: String,
}

fn main() {
    // describe the main window
    let main_window = WindowDesc::new(build_root_widget())
        .title(WINDOWS_TITLE)
        .window_size((800.0, 800.0));

    // create the initial app state
    let initial_state = HelloState {
        name: "VAL".into(),
    };

    // start the application
    AppLauncher::with_window(main_window)
        .launch(initial_state)
        .expect("launch failed");
}

fn build_root_widget() -> impl Widget<HelloState> {
    // a label that will determine its text based on the current app data.
    let label = Label::new(|data: &HelloState, _env: &_| format!("Hello, {}!", data.name));

    // a button that will increment the count when clicked.
    let button = Button::new("Click me!")
        .on_click(|_ctx, data: &mut HelloState, _env| data.name.push_str("!"));

    // arrange the two widgets vertically, with some padding
    Flex::column()
        .with_child(label)
        .with_spacer(VERTICAL_WIDGET_SPACING)
        .with_child(button)
}