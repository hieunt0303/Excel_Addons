export function setLoading(show) {
    if (show) {
        document.getElementById("main_group").style.display = 'none'
        document.getElementById("loader").style.display = 'block'
    } else {
        document.getElementById("main_group").style.display = 'block'
        document.getElementById("loader").style.display = 'none'
    }
}
