export function get(url, APIKey) {
    fetch(url, {
        method: "POST",
        redirect: "follow", // manual, *follow, error
        mode: "cors", // no-cors, *cors, same-origin
        cache: "default", // *default, no-cache, reload, force-cache, only-if-cached
        credentials: "same-origin",
        headers: {
            'Content-Type': 'application/json'
            // 'Content-Type': 'application/x-www-form-urlencoded',
        },
        body: JSON.stringify(APIKey)
    })
    .then(function(respond){
        return respond.json()
    })
    .then(function(dataOutput){
        console.log(dataOutput)
    })
    .catch(function(error){
        console.log(error)
    })
}