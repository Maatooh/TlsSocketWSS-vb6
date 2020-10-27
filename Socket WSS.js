
const socket1 = new WebSocket("wss://retromyths.com:5880");

socket1.onopen = () => {
    console.log("connected")
}

socket1.onmessage = function (event) {
    socket1.send(event.data);
    console.log(event.data);
  }
