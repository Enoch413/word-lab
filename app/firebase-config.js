(function () {
  const firebaseConfig = {
    apiKey: "AIzaSyA_6bQlCMYkYqkdqej8h57GGDWTPWXNml8",
    authDomain: "code-lab-2584c.firebaseapp.com",
    projectId: "code-lab-2584c",
    storageBucket: "code-lab-2584c.firebasestorage.app",
    messagingSenderId: "1073861558403",
    appId: "1:1073861558403:web:33d663a5ce7e0cadbddbb6",
    measurementId: "G-0LFRCLCR2Q"
  };

  [
    "__FIREBASE_CONFIG__",
    "FIREBASE_CONFIG",
    "firebaseConfig",
    "__firebaseConfig",
    "__CODELAB_FIREBASE_CONFIG__",
    "__WORD_LAB_FIREBASE_CONFIG__"
  ].forEach(function (key) {
    if (typeof window[key] === "undefined" || window[key] === null) {
      window[key] = firebaseConfig;
    }
  });
})();
