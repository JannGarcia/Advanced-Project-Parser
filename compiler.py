import os
import subprocess
import shutil
import stat
import multiprocessing
from multiprocessing import Process, Manager

MY_APPS_PATH = "D:\Workspaces\VSCode\of_v0.11.2_msys2_mingw32_release\\apps\myApps"

def clone_project(url: str) -> str:
    name = url.split("/")[-1].replace(".git", "")
    old_directory = os.getcwd()

    os.chdir(MY_APPS_PATH)

    if os.path.isdir(name):
        os.chdir(old_directory)
        return f"The directory {name} already exists"

    output = subprocess.run(["git", "clone", url], capture_output=True, text=True)
    os.chdir(old_directory)

    if output.returncode != 0:
        return f"Error cloning the repository {name}: {output.stderr}"
    return f"Repository {url} cloned successfully"


def compile_makefile_project(url: str) -> str:
    directory = os.path.join(MY_APPS_PATH, url)
    makefile_path = os.path.join(directory, "Makefile")

    if not os.path.isfile(makefile_path):
        return f"Makefile not found in {directory}"

    current_directory = os.getcwd()
    os.chdir(directory)

    output = subprocess.run(["make", "clean"], capture_output=True, text=True)
    if output.returncode != 0:
        os.chdir(current_directory)
        return f"Error cleaning the project {directory}: {output.stderr}"

    output = subprocess.run(["make", "-j6"], capture_output=True, text=True)
    os.chdir(current_directory)

    if output.returncode != 0:
        return f"Error compiling the project {directory}: {output.stderr}"

    return f"Project {directory} compiled successfully"


def on_rm_error(func, path, exc_info):
    os.chmod(path, stat.S_IWRITE)  # Change the file to writable
    func(path)  # Retry the removal

def delete_project(url: str):
    directory = os.path.join(MY_APPS_PATH, url)
    try:
        shutil.rmtree(directory, onerror=on_rm_error)
        return f"Deleted project directory {directory}."
    except Exception as e:
        return f"Failed to delete {directory}: {str(e)}"


def process_project(url: str, results: dict):
    name = url.split("/")[-1].replace(".git", "")
    project_directory = os.path.join(".", name)

    clone_result = clone_project(url)
    results[url] = [clone_result]  # Initialize results[url] with the clone result

    print("Processing:", url)

    if "cloned successfully" in clone_result:
        compile_result = compile_makefile_project(name)
        results[url] = results[url] + [compile_result]

        delete_result = delete_project(name)
        results[url] = results[url] + [delete_result]

    else:
        results[url] = results[url] + [f"Skipping compilation and deletion for {url} due to clone error."]



def worker(repos_list, results):
    for url in repos_list:
        process_project(url, results)

def compile_projects(repositories):
    manager = Manager()
    results = manager.dict()

    # Determine number of CPUs and divide repositories among them
    cpu_count = multiprocessing.cpu_count()
    chunk_size = (len(repositories) + cpu_count - 1) // cpu_count  # Calculate the chunk size per worker
    
    processes = []
    for i in range(cpu_count):
        # Assign each process its own chunk of repositories
        start_idx = i * chunk_size
        end_idx = min((i + 1) * chunk_size, len(repositories))
        repos_chunk = repositories[start_idx:end_idx]

        if repos_chunk:  # Check if there are any repos to assign
            p = Process(target=worker, args=(repos_chunk, results))
            p.start()
            processes.append(p)

    # Wait for all processes to finish
    for p in processes:
        p.join()

    return dict(results)


if __name__ == "__main__":
    file_path = "repos.txt"
    lines = []
    with open(file_path, "r") as file:
        lines = file.read().splitlines()

    compile_projects(lines)
