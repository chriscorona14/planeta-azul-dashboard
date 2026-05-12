import js from "@eslint/js";
export default [
    js.configs.recommended,
    {
        files: ["*.js"],
        languageOptions: {
            ecmaVersion: 2022,
            sourceType: "module",
            globals: {
                console: "readonly",
                Math: "readonly",
                Object: "readonly",
                String: "readonly",
                parseFloat: "readonly",
                isNaN: "readonly",
                Set: "readonly",
                Date: "readonly"
            }
        },
        rules: {
            "no-undef": "error",
            "no-unused-vars": "warn"
        }
    }
];
