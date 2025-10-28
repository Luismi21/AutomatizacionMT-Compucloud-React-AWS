# To learn more about how to use Nix to configure your environment
# see: https://developers.google.com/idx/guides/customize-idx-env
{ pkgs, ... }: {
  # Which nixpkgs channel to use.
  channel = "stable-24.05"; # or "unstable"

  # A list of packages to install, e.g. pkgs.go
  packages = [
    pkgs.nodejs_20
  ];

  # Sets environment variables in the workspace
  env = {};

  idx = {
    # VS Code extensions to install
    extensions = [
      "dbaeumer.vscode-eslint"
    ];

    # Enable previews and define a preview for the web server
    previews = {
      enable = true;
      previews = {
        web = {
          command = ["sh" "-c" "cd MT && npm run dev -- --port $PORT"];
          manager = "web";
        };
      };
    };

    # Workspace lifecycle hooks
    workspace = {
      # Runs when a workspace is first created
      onCreate = {
        npm-install = "cd MT && npm install";
      };

      # Runs when the workspace is (re)started
      onStart = {};
    };
  };
}
